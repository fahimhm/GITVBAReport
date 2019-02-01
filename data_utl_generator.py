# block-1: import library
#%%
import pandas as pd
import xlrd
import numpy as np

# block-2: load data source
#%%
# data_1 = pd.read_excel(r'Documents\DATA UTILITY.xlsx', sheet_name = 'DATA1')
# data_2 = pd.read_excel(r'Documents\DATA UTILITY.xlsx', sheet_name = 'DATA2')
data_1 = pd.read_excel(r'D:\14-Report Engineering\DATA UTILITY.xlsx', sheet_name = 'DATA1')
data_2 = pd.read_excel(r'D:\14-Report Engineering\DATA UTILITY.xlsx', sheet_name = 'DATA2')
data_3 = pd.read_excel(r'D:\14-Report Engineering\DATA UTILITY.xlsx', sheet_name = 'DATA3')

# block-3: merge two data
#%%
data_mh = pd.concat([data_1, data_2], ignore_index=True)
data_mh = data_mh.sort_values(['name', 'date_start'], ascending=True)
data_mh.reset_index(inplace=True, drop=True)

# block-3a: cek overlap
#%%
data_mh['remark'] = 'no_overlap'
for idx in range(1, len(data_mh.index)-1):
    if ((data_mh.loc[idx, 'name'] == data_mh.loc[idx+1, 'name']) & (data_mh.loc[idx+1, 'date_start'] < data_mh.loc[idx, 'date_finish'])) or ((data_mh.loc[idx, 'name'] == data_mh.loc[idx-1, 'name']) & (data_mh.loc[idx, 'date_start'] < data_mh.loc[idx-1, 'date_finish'])):
        data_mh.at[idx, 'remark'] = 'overlap'
data_mh[data_mh['remark'] == 'overlap']

# block-4: add Planned_spare_time for idle time in early shift
#%%
data_mh = data_mh.drop('remark', axis=1)
data_mh['only_date_start'] = data_mh['date_start'].dt.date
data_mh['hour_start'] = data_mh['date_start'].dt.hour + (data_mh['date_start'].dt.minute / 60)
col = data_mh.columns
data = []
for idx in range(1, len(data_mh.index)):
    if data_mh.loc[idx, 'day_cat'] == 'normal':
        if (data_mh.loc[idx-1, 'only_date_start'] != data_mh.loc[idx, 'only_date_start'] and data_mh.loc[idx, 'shift'] == 1) and data_mh.loc[idx, 'hour_start'] > 6.00:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 06:00:00'), data_mh.loc[idx, 'date_start'], 'NaN', 'NaN', 'NaN'])
        elif (data_mh.loc[idx-1, 'only_date_start'] != data_mh.loc[idx, 'only_date_start'] and data_mh.loc[idx, 'shift'] == 2) and data_mh.loc[idx, 'hour_start'] > 14.0:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 14:00:00'), data_mh.loc[idx, 'date_start'], 'NaN', 'NaN', 'NaN'])
        elif (data_mh.loc[idx, 'date_start'] - data_mh.loc[idx-1, 'date_finish']) / np.timedelta64(1, 'h') > 5 and data_mh.loc[idx, 'shift'] == 3 and data_mh.loc[idx, 'hour_start'] > 22.0:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 22:00:00'), data_mh.loc[idx, 'date_start'], 'NaN', 'NaN', 'NaN'])
        elif (data_mh.loc[idx-1, 'only_date_start'] != data_mh.loc[idx, 'only_date_start'] and data_mh.loc[idx, 'shift'] == 4) and data_mh.loc[idx, 'hour_start'] > 7.5:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 07:30:00'), data_mh.loc[idx, 'date_start'], 'NaN', 'NaN', 'NaN'])
    elif data_mh.loc[idx, 'day_cat'] == 'overtime':
        if (data_mh.loc[idx-1, 'only_date_start'] != data_mh.loc[idx, 'only_date_start'] and data_mh.loc[idx, 'shift'] == 1) and data_mh.loc[idx, 'hour_start'] > 6.00:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 06:00:00'), data_mh.loc[idx, 'date_start'], 'NaN', 'NaN', 'NaN'])
        elif (data_mh.loc[idx-1, 'only_date_start'] != data_mh.loc[idx, 'only_date_start'] and data_mh.loc[idx, 'shift'] == 2) and data_mh.loc[idx, 'hour_start'] > 13.5:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 13:30:00'), data_mh.loc[idx, 'date_start'], 'NaN', 'NaN', 'NaN'])
        elif (data_mh.loc[idx, 'date_start'] - data_mh.loc[idx-1, 'date_finish']) / np.timedelta64(1, 'h') > 5 and data_mh.loc[idx, 'shift'] == 3 and data_mh.loc[idx, 'hour_start'] > 21.0:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 21:00:00'), data_mh.loc[idx, 'date_start'], 'NaN', 'NaN', 'NaN'])
        elif (data_mh.loc[idx-1, 'only_date_start'] != data_mh.loc[idx, 'only_date_start'] and data_mh.loc[idx, 'shift'] == 4) and data_mh.loc[idx, 'hour_start'] > 7.0:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 07:00:00'), data_mh.loc[idx, 'date_start'], 'NaN', 'NaN', 'NaN'])
        elif (data_mh.loc[idx-1, 'only_date_start'] != data_mh.loc[idx, 'only_date_start'] and data_mh.loc[idx, 'shift'] == 5) and data_mh.loc[idx, 'hour_start'] > 12.0:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 12:00:00'), data_mh.loc[idx, 'date_start'], 'NaN', 'NaN', 'NaN'])
data_df = pd.DataFrame(data, columns=col)
data_df['date_start'] = pd.to_datetime(data_df['date_start'])
data_df['duration'] = (data_df['date_finish'] - data_df['date_start']) / np.timedelta64(1, 'h')
data_df['only_date_start'] = data_df['date_start'].dt.date
data_df['hour_start'] = data_df['date_start'].dt.hour + (data_df['date_start'].dt.minute / 60)
data_mh = pd.concat([data_mh, data_df], ignore_index=True)
data_mh = data_mh.sort_values(['name', 'date_start'], ascending=True)
data_mh.reset_index(inplace=True, drop=True)
data_mh = data_mh.drop(['only_date_start','hour_start'], axis=1)
# data_mh.head(15)

# block-5: add Planned_spare_time for idle time in end of shift
#%%
data_mh['only_date_finish'] = data_mh['date_finish'].dt.date
data_mh['hour_finish'] = data_mh['date_finish'].dt.hour + (data_mh['date_finish'].dt.minute / 60)
col = data_mh.columns
data = []
for idx in range(len(data_mh.index) - 1):
    if data_mh.loc[idx, 'day_cat'] == 'normal':
        if (data_mh.loc[idx+1, 'only_date_finish'] != data_mh.loc[idx, 'only_date_finish'] and data_mh.loc[idx, 'shift'] == 1) and data_mh.loc[idx, 'hour_finish'] < 14.5:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh.loc[idx, 'date_finish'], data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 14:30:00'), 'NaN', 'NaN', 'NaN'])
        elif data_mh.loc[idx+1, 'only_date_finish'] != data_mh.loc[idx, 'only_date_finish'] and data_mh.loc[idx, 'shift'] == 2 and data_mh.loc[idx, 'hour_finish'] < 22.5:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh.loc[idx, 'date_finish'], data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 22:30:00'), 'NaN', 'NaN', 'NaN'])
        elif (data_mh.loc[idx+1, 'date_start'] - data_mh.loc[idx, 'date_finish']) / np.timedelta64(1, 'h') > 5 and data_mh.loc[idx, 'shift'] == 3 and data_mh.loc[idx, 'hour_finish'] < 6.5:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh.loc[idx, 'date_finish'], data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 06:30:00'), 'NaN', 'NaN', 'NaN'])
        elif data_mh.loc[idx+1, 'only_date_finish'] != data_mh.loc[idx, 'only_date_finish'] and data_mh.loc[idx, 'shift'] == 4 and data_mh.loc[idx, 'hour_finish'] < 16.5:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh.loc[idx, 'date_finish'], data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 16:30:00'), 'NaN', 'NaN', 'NaN'])
    elif data_mh.loc[idx, 'day_cat'] == 'overtime':
        if data_mh.loc[idx+1, 'only_date_finish'] != data_mh.loc[idx, 'only_date_finish'] and data_mh.loc[idx, 'shift'] == 1 and data_mh.loc[idx, 'hour_finish'] < 13.0:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh.loc[idx, 'date_finish'], data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 13:00:00'), 'NaN', 'NaN', 'NaN'])
        elif data_mh.loc[idx+1, 'only_date_finish'] != data_mh.loc[idx, 'only_date_finish'] and data_mh.loc[idx, 'shift'] == 2 and data_mh.loc[idx, 'hour_finish'] < 21.0:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh.loc[idx, 'date_finish'], data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 21:00:00'), 'NaN', 'NaN', 'NaN'])
        elif (data_mh.loc[idx+1, 'date_start'] - data_mh.loc[idx, 'date_finish']) / np.timedelta64(1, 'h') > 5 and data_mh.loc[idx, 'shift'] == 3 and data_mh.loc[idx, 'hour_finish'] < 4.5:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh.loc[idx, 'date_finish'], data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 04:30:00'), 'NaN', 'NaN', 'NaN'])
        elif data_mh.loc[idx+1, 'only_date_finish'] != data_mh.loc[idx, 'only_date_finish'] and data_mh.loc[idx, 'shift'] == 4 and data_mh.loc[idx, 'hour_finish'] < 14.5:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh.loc[idx, 'date_finish'], data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 14:30:00'), 'NaN', 'NaN', 'NaN'])
        elif data_mh.loc[idx+1, 'only_date_finish'] != data_mh.loc[idx, 'only_date_finish'] and data_mh.loc[idx, 'shift'] == 5 and data_mh.loc[idx, 'hour_finish'] < 19.5:
            data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh.loc[idx, 'date_finish'], data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 19:30:00'), 'NaN', 'NaN', 'NaN'])
data_df = pd.DataFrame(data, columns=col)
data_df['date_finish'] = pd.to_datetime(data_df['date_finish'])
data_df['duration'] = (data_df['date_finish'] - data_df['date_start']) / np.timedelta64(1, 'h')
data_df['only_date_finish'] = data_df['date_finish'].dt.date
data_df['hour_finish'] = data_df['date_finish'].dt.hour + (data_df['date_finish'].dt.minute / 60)
data_mh = pd.concat([data_mh, data_df], ignore_index=True)
data_mh = data_mh.sort_values(['name', 'date_start'], ascending=True)
data_mh.reset_index(inplace=True, drop=True)
data_mh = data_mh.drop(['only_date_finish', 'hour_finish'], axis=1)
# data_mh.head(15)

# block-6: add Planned_spare_time in between activities
#%%
col = data_mh.columns
data = []
for idx in range(len(data_mh.index)-1):
    if (data_mh.loc[idx, 'date_finish'] != data_mh.loc[idx+1, 'date_start']) and (data_mh.loc[idx+1, 'date_start'] - data_mh.loc[idx, 'date_finish']) / np.timedelta64(1, 'h') < 5 and data_mh.loc[idx, 'name'] == data_mh.loc[idx+1, 'name']:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Planned', 'NaN', 'NaN', 'NaN', 'Planned_spare_time', data_mh.loc[idx, 'date_finish'], data_mh.loc[idx+1, 'date_start'], 'NaN'])
data = pd.DataFrame(data, columns=col)
data['duration'] = (data['date_finish'] - data['date_start']) / np.timedelta64(1, 'h')
data_mh = pd.concat([data, data_mh], ignore_index=True)
data_mh = data_mh.sort_values(['name', 'date_start'], ascending=True)
data_mh.reset_index(inplace=True, drop=True)
# data_mh.head(15)

# block-7: add break in each shift
#%%
data_mh['hour_start'] = data_mh['date_start'].dt.hour + (data_mh['date_start'].dt.minute / 60)
data_mh['hour_finish'] = data_mh['date_finish'].dt.hour + (data_mh['date_finish'].dt.minute / 60)
col = data_mh.columns
data = []
for idx in range(len(data_mh.index) - 1):
    if data_mh.loc[idx, 'shift'] == 1 and (data_mh.loc[idx, 'hour_start'] <= 10.0 and data_mh.loc[idx, 'hour_finish'] >= 10.5):
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 10:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 10:30:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'wo_status'], data_mh.loc[idx, 'system'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 10:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 10:00:00')
    elif data_mh.loc[idx, 'shift'] == 2 and data_mh.loc[idx, 'hour_start'] <= 18.0 and data_mh.loc[idx, 'hour_finish'] >= 18.5:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:30:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'wo_status'], data_mh.loc[idx, 'system'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 18:00:00')
    elif data_mh.loc[idx, 'day_cat'] == 'normal' and data_mh.loc[idx, 'shift'] == 3 and data_mh.loc[idx, 'hour_start'] <= 2.0 and data_mh.loc[idx, 'hour_finish'] >= 2.5:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 02:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 02:30:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'wo_status'], data_mh.loc[idx, 'system'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 02:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 02:00:00')
    elif data_mh.loc[idx, 'day_cat'] == 'normal' and data_mh.loc[idx, 'shift'] == 4 and data_mh.loc[idx, 'hour_start'] <= 12.0 and data_mh.loc[idx, 'hour_finish'] >= 13.0:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 12:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 13:00:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'wo_status'], data_mh.loc[idx, 'system'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 13:00:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 12:00:00')
    elif data_mh.loc[idx, 'day_cat'] == 'overtime' and data_mh.loc[idx, 'shift'] == 3 and data_mh.loc[idx, 'hour_start'] <= 24.0 and data_mh.loc[idx, 'hour_finish'] >= 0.5:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 00:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 00:30:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'wo_status'], data_mh.loc[idx, 'system'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 00:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 00:00:00')
    elif data_mh.loc[idx, 'day_cat'] == 'overtime' and data_mh.loc[idx, 'shift'] == 4 and data_mh.loc[idx, 'hour_start'] <= 10.0 and data_mh.loc[idx, 'hour_finish'] >= 10.5:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 10:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 10:30:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'wo_status'], data_mh.loc[idx, 'system'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 10:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 10:00:00')
    elif data_mh.loc[idx, 'shift'] == 5 and data_mh.loc[idx, 'hour_start'] <= 18.0 and data_mh.loc[idx, 'hour_finish'] >= 18.5:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:30:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'wo_status'], data_mh.loc[idx, 'system'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 18:00:00')
data = pd.DataFrame(data, columns=col)
data['date_start'] = pd.to_datetime(data['date_start'])
data['date_finish'] = pd.to_datetime(data['date_finish'])
data['duration'] = (data['date_finish'] - data['date_start']) / np.timedelta64(1, 'h')
data_mh = pd.concat([data, data_mh], ignore_index=True)
data_mh = data_mh.sort_values(['name', 'date_start'], ascending=True)
data_mh = data_mh.drop(['hour_start', 'hour_finish'], axis=1)
data_mh['duration'] = (data_mh['date_finish'] - data_mh['date_start']) / np.timedelta64(1, 'h')
data_mh = data_mh.drop(data_mh[data_mh.duration == 0].index, axis=0)
data_mh.reset_index(inplace=True, drop=True)

#%%
data_mh[data_mh['duration'] <= 0.0]

# block-8: convert Planned_spare_time to idle_time for Planned_spare_time is more than 2.5 hours
#%%
data_mh['flag'] = 0.0
sumif = 0.0
for idx in range(1, len(data_mh.index)):
    if data_mh.loc[idx, 'day'] == data_mh.loc[idx-1, 'day']:
        if data_mh.loc[idx, 'machine'] == 'Planned_spare_time':
            sumif = sumif + data_mh.loc[idx, 'duration']
            data_mh.at[idx, 'flag'] = sumif
    else:
        sumif = 0.0
        if data_mh.loc[idx, 'machine'] == 'Planned_spare_time':
            sumif = sumif + data_mh.loc[idx, 'duration']
            data_mh.at[idx, 'flag'] = sumif
for idx in data_mh.index:
    if data_mh.loc[idx, 'flag'] >= 2.5:
        data_mh.at[idx, 'task_cat'] = 'Unplanned'
        data_mh.at[idx, 'machine'] = 'idle_time'
data_mh = data_mh.drop('flag', axis=1)
data_mh['year'] = data_mh['year'].astype('category')
data_mh['week'] = data_mh['week'].astype('category')
data_mh['shift'] = data_mh['shift'].astype('category')

#%%
# ---- save to excel -----
writer = pd.ExcelWriter('data_utility.xlsx')
data_mh.to_excel(writer, 'DATA')
writer.save()

#%%
# ---- create barplot with bokeh----
# %Productivity report per month - 2018
from bokeh.io import show, output_notebook
from bokeh.models import ColumnDataSource, FactorRange, Span
from bokeh.plotting import figure
import itertools

mh_2018 = data_mh[data_mh['year'] == 2018] # replace 2018 >> 2019
group_prod = mh_2018.groupby(['name', 'month']).apply(lambda x: x[(x['task_cat'] != 'Break') & (x['task_cat'] != 'Unplanned')]['duration'].sum()*100.0/x[x['task_cat'] != 'Break']['duration'].sum()).unstack()

names = group_prod.index.get_level_values(0).tolist()
months = data_mh[data_mh['year'] == 2018]['month'].unique().tolist() # replace 2018 >> 2019

x = [(name, str(month)) for name in names for month in months]
l = [group_prod.loc[name, :].tolist() for name in names]
l = list(itertools.chain(*l))
counts = tuple(l)

source = ColumnDataSource(data = dict(x=x, counts=counts))

output_notebook()
plot_mh = figure(x_range = FactorRange(*x), plot_height = 400, title = "%Productivity", toolbar_location = "right")
plot_mh.vbar(x = 'x', top = 'counts', width = 0.9, source = source)

plot_mh.y_range.start = 0
plot_mh.x_range.range_padding = 0.1
plot_mh.xaxis.major_label_orientation = 1
plot_mh.xgrid.grid_line_color = None

target_line = Span(location = 70, dimension = 'width', line_color = 'red', line_dash = 'dashed', line_width = 2)
plot_mh.add_layout(target_line)

show(plot_mh)

#%%
# %Productivity VS %WO-Realization 2018
data_3 = data_3[data_3['year'] == 2018]
x_month = [str(month) for month in months]
prod_month = mh_2018.groupby('month').apply(lambda x: x[(x['task_cat'] != 'Break') & (x['task_cat'] != 'Unplanned')]['duration'].sum()*100.0/x[x['task_cat'] != 'Break']['duration'].sum()).tolist()
wo_month = data_3.groupby('month').apply(lambda x: x[x['wo_status'] == 'done']['wo'].nunique()*100.0/x['wo'].nunique()).tolist()

output_notebook()
mh_vs_wo = figure(x_range = x_month, plot_height = 400, title = "%Productivity vs %WO realization", toolbar_location = "right")
mh_vs_wo.vbar(x = x_month, top = prod_month, width = 0.9)
mh_vs_wo.line(x = x_month, y = wo_month, color = "red", line_width = 2)
mh_vs_wo.circle(x = x_month, y = wo_month, fill_color = "white")

mh_vs_wo.xgrid.grid_line_color = None
mh_vs_wo.y_range.start = 0

show(mh_vs_wo)