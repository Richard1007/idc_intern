import pandas as pd
import numpy as np
import os
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

channelsum=pd.read_excel(r'C:\Users\mjia\Desktop\220614 pc slaein city fcst\channel_mapping.xlsx',sheet_name='channeluse')
aspuse=pd.read_excel(r'C:\Users\mjia\Desktop\220614 pc slaein city fcst\pc_fcst.xlsx',sheet_name='aspuse')
channelmix=pd.read_excel(r'C:\Users\mjia\Desktop\220614 pc slaein city fcst\channelmix.xlsx',sheet_name='Sheet5')
segpro=pd.read_excel(r'C:\Users\mjia\Desktop\220614 pc slaein city fcst\segpro_smb_tran3.xlsx',sheet_name='Sheet1')

print(channelsum.loc[channelsum.Unitschannelmapping==0,:].count())
print(segpro.loc[segpro.Units==0,:].count())

segpro1=segpro.loc[segpro.Units>0,:]
channelsum1=channelsum.loc[channelsum.Unitschannelmapping>0,:]

print(segpro1.Units.sum())
print(channelsum1.Unitschannelmapping.sum())
print(channelsum.Unitschannelmapping.sum())


channelsum1_total_p=channelsum1.groupby(['Quarter','Product','Segment']).agg({'Unitschannelmapping':'sum'})
channelsum1_total=pd.DataFrame(channelsum1_total_p.to_records())
channelsum1_total.rename(columns={'Unitschannelmapping':'Total_Unitschannelmapping'},inplace=True)
channelsum2=pd.merge(channelsum1,channelsum1_total,on=['Quarter','Product','Segment'],how='left')
channelsum2['channel_mix_channelmapping']=channelsum2['Unitschannelmapping']/channelsum2['Total_Unitschannelmapping']

channelsum2.count()

segpro1_total_p=segpro1.groupby(['Quarter','Product','Segment']).agg({'Units':'sum'})
segpro1_total=pd.DataFrame(segpro1_total_p.to_records())
segpro1_total.rename(columns={'Units':'Total_Units'},inplace=True)
segpro2=pd.merge(segpro1,segpro1_total,on=['Quarter','Product','Segment'],how='left')
segpro2['provmix']=segpro2['Units']/segpro2['Total_Units']


segpro2.drop(['Country','Forecast Version','Segment (Internal)'],inplace=True,axis=1)

segpro3=pd.merge(segpro2,channelsum2,on=['Quarter','Product','Segment'],how='outer')


channelmix1=pd.melt(channelmix,id_vars=['Quarter','Product','Segment','Province'],value_vars=['Direct - Inbound/Outbound','Direct - Internet',
                                        'Direct - Store','InDirect - Dealer/VAR/SI' ,'InDirect - eTailer','InDirect - LFR'  ,'InDirect - Retail','InDirect - Telco' ],
                   var_name='Channel',value_name='channelmix')

segpro4=pd.merge(segpro3,channelmix1,on=['Quarter','Product','Segment','Province','Channel'],how='left')

segpro5=segpro4[['Quarter','Product','Segment','Province','Channel','Total_Units','Units','Unitschannelmapping','provmix','channel_mix_channelmapping','channelmix']]


segpro5.rename(columns={'Units':'Unitsprovince'},inplace=True)

segpro5['channelmix_new']=(segpro5['channelmix']*3+segpro5['channel_mix_channelmapping'])/4
segpro5_total_p=segpro5.groupby(['Quarter','Product','Segment','Province']).agg({'channelmix_new':'sum'})
segpro5_total=pd.DataFrame(segpro5_total_p.to_records())
segpro5_total.rename(columns={'channelmix_new':'Total_channelmix_new'},inplace=True)
segpro6=pd.merge(segpro5,segpro5_total,on=['Quarter','Product','Segment','Province'],how='left')
segpro6['channel_mix_new1']=segpro6['channelmix_new']/segpro6['Total_channelmix_new']
print(segpro6['channel_mix_new1'].sum())
print(segpro6['channelmix_new'].sum())

segpro7=segpro6[['Quarter','Product','Segment','Province','Channel','Total_Units','Unitsprovince','Unitschannelmapping','provmix','channel_mix_new1']]
segpro7.rename(columns={'channel_mix_new1':'channelmix'},inplace=True)
segpro7['Units_channeluse0']=segpro7['Unitsprovince']*segpro7['channelmix']

#这一步定义非常重要
clist=['InDirect - Dealer/VAR/SI', 'InDirect - eTailer', 'InDirect - Retail', 'Direct - Inbound/Outbound', 'Direct - Internet','InDirect - LFR','Direct - Store', 'InDirect - Telco']
clist1=list(reversed(clist))
print(clist1)

segpro00=segpro7
print(segpro00.count())

chnnl=pd.DataFrame(columns=['Quarter','Product','Segment','Province','Channel','Units_channeluse_final'])

for i,x in enumerate(clist1[:7]):
    print(x)
    print(clist1[(i+1):])
    channel1=segpro00.loc[(segpro00['Channel']==x)]
    channel_other1=segpro00.loc[(segpro00['Channel'].isin(clist1[(i+1):])),:]
    channel_mapping1=channelsum1[(channelsum1['Channel']==x)]
  
    test_channel1_p=channel1.groupby(['Quarter','Product','Segment']).agg({'Units_channeluse0':'sum'})
    test_channel1=pd.DataFrame(test_channel1_p.to_records())
    test_channel1.rename(columns={'Units_channeluse0':'Unitschanneluse'},inplace=True)

    test_channel_other1_p=channel_other1.groupby(['Quarter','Product','Segment']).agg({'Units_channeluse0':'sum'})
    test_channel_other1=pd.DataFrame(test_channel_other1_p.to_records())
    test_channel_other1.rename(columns={'Units_channeluse0':'Unitschanneluse_other'},inplace=True)
    
    
    test_all=pd.merge(test_channel1,test_channel_other1,on=['Quarter','Product','Segment'],how='left')#这里好像要改20220720
    test_all.fillna(0)
    
    test_all3=pd.merge(channel1,test_all,on=['Quarter','Product','Segment'],how='left')
    
    
    test_all3['difference']=test_all3['Unitschannelmapping']-test_all3['Unitschanneluse']
    
    
    test_all3['check']=-test_all3['difference']+test_all3['Unitschanneluse_other']
    
    print('Haha1',x,test_all3.count())
    print('Haha1',x,test_all3.loc[test_all3.check<0,:].count())
    print('Haha1',x,test_all3.loc[test_all3.check<0,:].head(10))
    
    #test_all1=test_all.loc[test_all.check>=0,:]
    test_all3.loc[test_all3.check<0,'difference']=test_all3.loc[test_all3.check<0,'Unitschanneluse_other']
    
    #test_all2=pd.merge(test_all1,segpro_withmix,on=['Quarter','Product','Segment'],how='left')
    
    
    #test_all3=pd.merge(channel1,test_all1,on=['Quarter','Product','Segment','Province'],how='outer')
    #print('Haha2',x,test_all3.count())
    
    
    
    #test_all3['Units_channeluse_final']=test_all3['Units_channeluse0']+test_all3['difference']*test_all3['provmix']
    test_all3['provmix1']=test_all3['Units_channeluse0']/test_all3['Unitschanneluse']
    #test_all3.loc[test_all3.Unitschanneluse==0,'provmix1']=0
    test_all3['Units_channeluse_final']=test_all3['Units_channeluse0']+test_all3['difference']*test_all3['provmix1']
    test_all3['check1']=test_all3['provmix1']*test_all3['difference']
    
    #print('Haha3',test_all3.count())
    
    print('Haha4',x,test_all3.Units_channeluse_final.sum(),channel_mapping1.Unitschannelmapping.sum(),test_channel1.Unitschanneluse.sum())
    print('Haha4',x,(channel_mapping1.Unitschannelmapping.sum()-test_channel1.Unitschanneluse.sum()))
    test_all4=test_all3.drop(['channelmix','Units_channeluse0','Unitschanneluse','Unitschanneluse_other','Unitschannelmapping','difference','check','provmix','provmix1','Total_Units','Unitsprovince'],axis=1)
    
    #test_all4['Channel']=x
    
    test_all4_reorder=test_all4[['Quarter','Product','Segment','Province','Channel','Units_channeluse_final']]
    chnnl=pd.concat([chnnl,test_all4_reorder],axis=0)
    
    channel_other2=pd.merge(channel_other1,test_all3[['Quarter','Product','Segment','Province','difference']],on=['Quarter','Product','Segment','Province'],how='left')
    print('Haha5',x,channel_other2.count())
    channel_other2.difference.fillna(0,inplace=True)

   
    channel_other2_totalp=channel_other2.groupby(['Quarter','Product','Segment']).agg({'Units_channeluse0':'sum'})
    channel_other2_total=pd.DataFrame(channel_other2_totalp.to_records())
    channel_other2_total.rename(columns={'Units_channeluse0':'Total_Unitschanneluse'},inplace=True)    
    
    channel_other3=pd.merge(channel_other2,channel_other2_total,on=['Quarter','Product','Segment'],how='left')
    
    channel_other3['provmix1']=channel_other3['Units_channeluse0']/channel_other3['Total_Unitschanneluse']
    print('Haha6',x,channel_other3.count())
    
    channel_other3['Units_channeluse_final']=channel_other3['Units_channeluse0']-channel_other3['difference']*channel_other3['provmix1']
    print('Haha7',x,channel_other3.Units_channeluse_final.sum(),channel_other2_total.Total_Unitschanneluse.sum())
    print('Haha7',x,(channel_other3.Units_channeluse_final.sum()-channel_other2_total.Total_Unitschanneluse.sum()))
    
    channel_other3.drop(['difference','provmix1','Units_channeluse0','Total_Unitschanneluse'],axis=1,inplace=True)
    channel_other3.rename(columns={'Units_channeluse_final':'Units_channeluse0'},inplace=True)
    
    segpro00=channel_other3

print(test_all4.loc[test_all4.Units_channeluse_final<0,:].count())
print(chnnl.loc[chnnl.Units_channeluse_final<0,:].count())

print(test_all4.Channel.value_counts())
print(chnnl.Channel.value_counts())
segpro00.rename(columns={'Units_channeluse0':'Units_channeluse_final'},inplace=True)
final=pd.concat([segpro00[['Quarter','Product','Segment','Province','Channel','Units_channeluse_final']],chnnl],axis=0)

print(final.count())
print(final.groupby('Channel')['Units_channeluse_final'].sum())
print(final.Units_channeluse_final.sum())

print(channelsum1.groupby('Channel')['Unitschannelmapping'].sum())
print(channelsum1.Unitschannelmapping.sum())

final.to_excel(r'C:\Users\mjia\Desktop\220614 pc slaein city fcst\test_output_08302022.xlsx')
