import pandas as pd
import numpy as np
import os
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

channelsum=pd.read_excel(r'C:\Users\mjia\Desktop\220614\channel_mapping.xlsx',sheet_name='channeluse')
aspuse=pd.read_excel(r'C:\Users\mjia\Desktop\220614\pc_fcst.xlsx',sheet_name='aspuse')
channelmix=pd.read_excel(r'C:\Users\mjia\Desktop\220614\channelmix.xlsx',sheet_name='Sheet5')
segpro=pd.read_excel(r'C:\Users\mjia\Desktop\220614\segpro_smb_tran3.xlsx',sheet_name='Sheet1')


channelmix1=pd.melt(channelmix,id_vars=['Quarter','Product','Segment','Province'],value_vars=['Direct - Inbound/Outbound','Direct - Internet',
                                        'Direct - Store','InDirect - Dealer/VAR/SI' ,'InDirect - eTailer','InDirect - LFR'  ,'InDirect - Retail','InDirect - Telco' ],
                   var_name='Channel',value_name='channelmix')

segpro1=pd.merge(segpro,channelmix1,on=['Quarter','Product','Segment','Province'],how='left')
channelmix1.channelmix.sum()
segpro.count()
segpro1.channelmix.sum()
segpro1['Units_channeluse0']=segpro1['Units']*segpro1['channelmix']
segpro1.Units_channeluse0.sum()
segpro.Units.sum()
segpro1.drop(['Units'],axis=1,inplace=True)



test5_p=segpro.groupby(['Quarter','Product','Segment']).agg({'Units':'sum'})
test5=pd.DataFrame(test5_p.to_records())
test5.rename(columns={'Units':'Unitsprovincesum'},inplace=True)

segpro_withmix=pd.merge(segpro,test5,on=['Quarter','Product','Segment'],how='left')
segpro_withmix['provmix']=segpro_withmix['Units']/segpro_withmix['Unitsprovincesum']
segpro_withmix.drop(['Country','Forecast Version','Segment (Internal)','Units','Unitsprovincesum'],axis=1,inplace=True)
segpro_withmix.head(5)


segpro_withmix.count()
#这一步定义非常重要
clist=['InDirect - Dealer/VAR/SI', 'InDirect - eTailer', 'InDirect - Retail', 'Direct - Inbound/Outbound', 'Direct - Internet','InDirect - LFR','Direct - Store', 'InDirect - Telco']
clist1=list(reversed(clist))
print(clist1)

segpro00=segpro1
segpro00.count()

chnnl=segpro.loc[segpro.Country=='']


#h=0,1,2,3,4,5,6 从0至6分次运行，每次运行完毕需要调整 segpro00
for i,x in enumerate(clist1[h:(h+1)]):
    print(x)
    print(clist1[(h+1):])
    channel1=segpro00.loc[(segpro00['Channel']==x)]
    channel_other1=segpro00.loc[(segpro00['Channel'].isin(clist1[(h+1):])),:]
    channel_mapping1=channelsum[(channelsum['Channel']==x)]
  
    test_channel1_p=channel1.groupby(['Quarter','Product','Segment']).agg({'Units_channeluse0':'sum'})
    test_channel1=pd.DataFrame(test_channel1_p.to_records())
    test_channel1.rename(columns={'Units_channeluse0':'Unitschanneluse'},inplace=True)

    test_channel_other1_p=channel_other1.groupby(['Quarter','Product','Segment']).agg({'Units_channeluse0':'sum'})
    test_channel_other1=pd.DataFrame(test_channel_other1_p.to_records())
    test_channel_other1.rename(columns={'Units_channeluse0':'Unitschanneluse_other'},inplace=True)
    
    test_channel_mapping_p=channel_mapping1.groupby(['Quarter','Product','Segment']).agg({'Unitschannelmapping':'sum'})
    test_channel_mapping=pd.DataFrame(test_channel_mapping_p.to_records())
    test_channel_mapping.rename(columns={'Unitschanneluse0':'Unitschanneluse'},inplace=True)
    
    test_all0=pd.merge(test_channel1,test_channel_other1,on=['Quarter','Product','Segment'],how='outer')
    test_all=pd.merge(test_all0,test_channel_mapping,on=['Quarter','Product','Segment'],how='left')
    #test_all['Channel']=x
    
    test_all.loc[test_all.Unitschannelmapping.isnull(),'Unitschannelmapping']=0
    
    test_all['difference']=test_all['Unitschannelmapping']-test_all['Unitschanneluse']
    
    
    test_all['check']=-test_all['difference']+test_all['Unitschanneluse_other']
    
    print('Haha1',test_all.count())
    #print(test_all.loc[test_all.check<0,:].count())
    #print(test_all.loc[test_all.check<0,:].head(10))
    
    #test_all1=test_all.loc[test_all.check>=0,:]
    test_all1=test_all
    print('Haha2',test_all.loc[test_all.check<0,:])
    
    
    test_all2=pd.merge(test_all1,segpro_withmix,on=['Quarter','Product','Segment'],how='left')
    
    
    test_all3=pd.merge(channel1,test_all2,on=['Quarter','Product','Segment','Province'],how='inner')

    
    #test_all3['Units_channeluse_final']=test_all3['Units_channeluse0']+test_all3['difference']*test_all3['provmix']
    test_all3['provmix1']=test_all3['Units_channeluse0']/test_all3['Unitschanneluse']
    test_all3.loc[test_all3.Unitschanneluse==0,'provmix1']=0
    test_all3['Units_channeluse_final']=test_all3['Units_channeluse0']+test_all3['difference']*test_all3['provmix1']
    test_all3['check1']=test_all3['provmix1']*test_all3['difference']
    
    #print('Haha3',test_all3.count())
    
    print('Haha4',test_all3.Units_channeluse_final.sum(),channel_mapping1.Unitschannelmapping.sum())

    test_all4=test_all3.drop(['channelmix','Units_channeluse0','Unitschanneluse','Unitschanneluse_other','Unitschannelmapping','difference','check','provmix','provmix1'],axis=1)
    
    channel_other2=pd.merge(channel_other1,test_all1,on=['Quarter','Product','Segment'],how='outer')
    print('Haha5',channel_other2.count())
    
    channel_other3=channel_other2.loc[channel_other2.check.notnull(),:]
    print('Haha6',channel_other3.count())
    
    
    channel_other4=pd.merge(channel_other3,test_all3.loc[:,['Quarter','Product','Segment','Province','provmix1','check1']],on=['Quarter','Product','Segment','Province'],how='outer')
    print('Haha7',channel_other4.count())
    
    channel_other4.groupby(['Product','Segment','Province']).agg({'Units_channeluse0':'sum'})

    test_channel_other4_p=channel_other4.groupby(['Quarter','Product','Segment','Province']).agg({'Units_channeluse0':'sum'})
    test_channel_other4=pd.DataFrame(test_channel_other4_p.to_records())
    test_channel_other4.rename(columns={'Units_channeluse0':'Units_channeluse_again'},inplace=True)
    


    channel_other5=pd.merge(channel_other4,test_channel_other4,on=['Quarter','Product','Segment','Province'],how='left')
    print('!!!',channel_other5.loc[channel_other5.check1>channel_other5.Units_channeluse_again,:].count())
    channel_other5.rename(columns={'Units_channeluse0':'Units_channeluse_00','channelmix':'channelmix_00'},inplace=True)
    channel_other5['channelmix']=channel_other5['Units_channeluse_00']/channel_other5['Units_channeluse_again']
    channel_other5['Units_channeluse0']=channel_other5['Units_channeluse_00']-channel_other5['difference']*channel_other5['provmix1']*channel_other5['channelmix']

    #print('Haha8',channel_other5.loc[channel_other5.Units_channeluse0<0,:].count())

    channel_other6=channel_other5.drop(['channelmix_00','Units_channeluse_00','Unitschanneluse','Unitschanneluse_other','Unitschannelmapping','difference','check','check1','provmix1','Units_channeluse_again'],axis=1)

    segpro00=channel_other6
    
    test_all4['Channel']=x
    chnnl=pd.concat([chnnl,test_all4],axis=0)
    
    
print(segpro00.count())
#segpro00.loc[segpro00.channelmix.isnull(),:].head(100)
#segpro00.loc[segpro00.channelmix.isnull(),:].Segment.value_counts()
#segpro00.loc[segpro00.Units_channeluse0.isnull(),'channelmix']=0
#segpro00.loc[segpro00.Units_channeluse0.isnull(),'Units_channeluse0']=0

test_all.loc[test_all.difference.isnull(),:].head(5)
#segpro00.loc[segpro00.Units_channeluse0<0,:].head(100)
#segpro00.loc[segpro00.Units_channeluse0<0,'Units_channeluse0']=-segpro00.loc[segpro00.Units_channeluse0<0,'Units_channeluse0']

channel_other5.loc[channel_other5.check1>channel_other5.Units_channeluse_again,:]

chnnl.count()
chnnl.Channel.value_counts()
chnnl.loc[chnnl.Units_channeluse_final<0,:]

chnnl.to_excel(r'C:\Users\mjia\Desktop\220614\test_output_reverse1.xlsx')
channel_other5.to_excel(r'C:\Users\mjia\Desktop\220614\test_output_reverse2.xlsx') 


##进行value生成
output=pd.read_excel(r'C:\Users\mjia\Desktop\220614\test_output_reverse2.xlsx',sheet_name='Sheet3')
output.Units_channeluse_final.sum()
output1=pd.merge(output,aspuse,on=['Quarter','Product','Segment'],how='left')
output1.to_excel(r'C:\Users\mjia\Desktop\220614\test_reverse_output3.xlsx')