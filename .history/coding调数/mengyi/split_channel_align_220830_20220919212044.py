import pandas as pd
import numpy as np
import os
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

channelsum=pd.read_excel('/Users/richardpang/Desktop/idc_intern-1/coding调数/mengyi/channel_mapping.xlsx',sheet_name='channeluse')
aspuse=pd.read_excel('/Users/richardpang/Desktop/idc_intern-1/coding调数/mengyi/pc_fcst.xlsx',sheet_name='aspuse')
channelmix=pd.read_excel('/Users/richardpang/Desktop/idc_intern-1/coding调数/mengyi/channelmix.xlsx',sheet_name='Sheet5')
segpro=pd.read_excel('/Users/richardpang/Desktop/idc_intern-1/coding调数/mengyi/segpro_smb_tran3.xlsx',sheet_name='Sheet1')

### STEP1: Merge
def merge(channelsum,segpro,merge_label,drop_label):
    channelsum1=channelsum.loc[channelsum.Unitschannelmapping>0,:]
    channelsum1_total_p=channelsum1.groupby(merge_label).agg({'Unitschannelmapping':'sum'})
    channelsum1_total=pd.DataFrame(channelsum1_total_p.to_records())
    channelsum1_total.rename(columns={'Unitschannelmapping':'Total_Unitschannelmapping'},inplace=True)
    channelsum2=pd.merge(channelsum1,channelsum1_total,on=merge_label,how='left')
    channelsum2['channel_mix_channelmapping']=channelsum2['Unitschannelmapping']/channelsum2['Total_Unitschannelmapping']
    channelsum2.count()
    segpro1=segpro.loc[segpro.Units>0,:]
    segpro1_total_p=segpro1.groupby(merge_label).agg({'Units':'sum'})
    segpro1_total=pd.DataFrame(segpro1_total_p.to_records())
    segpro1_total.rename(columns={'Units':'Total_Units'},inplace=True)
    segpro2=pd.merge(segpro1,segpro1_total,on=merge_label,how='left')
    segpro2['provmix']=segpro2['Units']/segpro2['Total_Units']
    segpro2.drop(drop_label,inplace=True,axis=1)
    segpro3=pd.merge(segpro2,channelsum2,on=merge_label,how='outer')
    return channelsum1,segpro3

merge_label = ['Quarter','Product','Segment']
drop_label = ['Country','Forecast Version','Segment (Internal)']  

channelsum1,segpro3 = merge(channelsum,segpro,merge_label,drop_label)



### STEP2: Melt
def melt(channelmix,melt_id_vars,melt_value_vars,melt_var_name,melt_value_name):
    channelmix1=pd.melt(channelmix,id_vars=melt_id_vars,value_vars=melt_value_vars,var_name=melt_var_name,value_name=melt_value_name)
    return channelmix1

melt_id_vars = ['Quarter','Product','Segment','Province']
melt_value_vars = ['Direct - Inbound/Outbound','Direct - Internet','Direct - Store','InDirect - Dealer/VAR/SI' ,'InDirect - eTailer','InDirect - LFR'  ,'InDirect - Retail','InDirect - Telco']
melt_var_name = 'Channel'
melt_value_name='channelmix'

channelmix1 = melt(channelmix,melt_id_vars,melt_value_vars,melt_var_name,melt_value_name)



### STEP3: Merge again
def merge_again(segpro3,channelmix1,merge_label):
    segpro4=pd.merge(segpro3,channelmix1,on=merge_label,how='left')
    segpro5=segpro4[['Quarter','Product','Segment','Province','Channel','Total_Units','Units','Unitschannelmapping','provmix','channel_mix_channelmapping','channelmix']]
    segpro5.rename(columns={'Units':'Unitsprovince'},inplace=True)
    segpro5['channelmix_new']=(segpro5['channelmix']*3+segpro5['channel_mix_channelmapping'])/4
    segpro5_total_p=segpro5.groupby(['Quarter','Product','Segment','Province']).agg({'channelmix_new':'sum'})
    segpro5_total=pd.DataFrame(segpro5_total_p.to_records())
    segpro5_total.rename(columns={'channelmix_new':'Total_channelmix_new'},inplace=True)
    segpro6=pd.merge(segpro5,segpro5_total,on=['Quarter','Product','Segment','Province'],how='left')
    segpro6['channel_mix_new1']=segpro6['channelmix_new']/segpro6['Total_channelmix_new']
    segpro7=segpro6[['Quarter','Product','Segment','Province','Channel','Total_Units','Unitsprovince','Unitschannelmapping','provmix','channel_mix_new1']]
    segpro7.rename(columns={'channel_mix_new1':'channelmix'},inplace=True)
    segpro7['Units_channeluse0']=segpro7['Unitsprovince']*segpro7['channelmix']
    return segpro7

merge_label = ['Quarter','Product','Segment','Province','Channel']
segpro7 = merge_again(segpro3,channelmix1,merge_label)





### STEP4:Distribute
def distribute(clist,segpro7,chnnl_columns):
    clist1=list(reversed(clist))
    segpro00=segpro7
    chnnl=pd.DataFrame(columns=chnnl_columns)
    for i,x in enumerate(clist1[:7]):
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
        test_all3.loc[test_all3.check<0,'difference']=test_all3.loc[test_all3.check<0,'Unitschanneluse_other']
        test_all3['provmix1']=test_all3['Units_channeluse0']/test_all3['Unitschanneluse']
        test_all3['Units_channeluse_final']=test_all3['Units_channeluse0']+test_all3['difference']*test_all3['provmix1']
        test_all3['check1']=test_all3['provmix1']*test_all3['difference']
        test_all4=test_all3.drop(['channelmix','Units_channeluse0','Unitschanneluse','Unitschanneluse_other','Unitschannelmapping','difference','check','provmix','provmix1','Total_Units','Unitsprovince'],axis=1)
        test_all4_reorder=test_all4[['Quarter','Product','Segment','Province','Channel','Units_channeluse_final']]
        chnnl=pd.concat([chnnl,test_all4_reorder],axis=0)
        channel_other2=pd.merge(channel_other1,test_all3[['Quarter','Product','Segment','Province','difference']],on=['Quarter','Product','Segment','Province'],how='left')
        channel_other2.difference.fillna(0,inplace=True)
        channel_other2_totalp=channel_other2.groupby(['Quarter','Product','Segment']).agg({'Units_channeluse0':'sum'})
        channel_other2_total=pd.DataFrame(channel_other2_totalp.to_records())
        channel_other2_total.rename(columns={'Units_channeluse0':'Total_Unitschanneluse'},inplace=True)    
        channel_other3=pd.merge(channel_other2,channel_other2_total,on=['Quarter','Product','Segment'],how='left')
        channel_other3['provmix1']=channel_other3['Units_channeluse0']/channel_other3['Total_Unitschanneluse']
        channel_other3['Units_channeluse_final']=channel_other3['Units_channeluse0']-channel_other3['difference']*channel_other3['provmix1']
        channel_other3.drop(['difference','provmix1','Units_channeluse0','Total_Unitschanneluse'],axis=1,inplace=True)
        channel_other3.rename(columns={'Units_channeluse_final':'Units_channeluse0'},inplace=True)
        segpro00=channel_other3
        return segpro00



clist=['InDirect - Dealer/VAR/SI', 'InDirect - eTailer', 'InDirect - Retail', 'Direct - Inbound/Outbound', 'Direct - Internet','InDirect - LFR','Direct - Store', 'InDirect - Telco']
chnnl_columns=['Quarter','Product','Segment','Province','Channel','Units_channeluse_final']
segpro00 = distribute(clist,segpro7,chnnl_columns)


segpro00.rename(columns={'Units_channeluse0':'Units_channeluse_final'},inplace=True)
final=pd.concat([segpro00[['Quarter','Product','Segment','Province','Channel','Units_channeluse_final']],chnnl],axis=0)
final.to_excel('/Users/richardpang/Desktop/idc_intern-1/coding调数/mengyi/test_output_08302022.xlsx')
