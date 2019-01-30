# block-1: import library
#%%
import pandas as pd
import xlrd
import numpy as np

# block-2: load data source
#%%
# data_1 = pd.read_excel(r'C:\Users\Fahim Hadi Maula\Git\GITVBAReport\DATA UTILITY.xlsx', sheet_name = 'DATA1')
# data_2 = pd.read_excel(r'C:\Users\Fahim Hadi Maula\Git\GITVBAReport\DATA UTILITY.xlsx', sheet_name = 'DATA2')
data_1 = pd.read_excel(r'C:\Users\maula.fahim\github\dept_report\DATA UTILITY.xlsx', sheet_name = 'DATA1')
data_2 = pd.read_excel(r'C:\Users\maula.fahim\github\dept_report\DATA UTILITY.xlsx', sheet_name = 'DATA2')

# block-3: merge two data
#%%
data_mh = pd.concat([data_1, data_2], ignore_index=True)
data_mh = data_mh.sort_values(['name', 'date_start'], ascending=True)
data_mh.reset_index(inplace=True, drop=True)
# data_mh.head(15)

# block-4: add Planned_spare_time for idle time in early shift
#%%
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
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 10:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 10:00:00')
    elif data_mh.loc[idx, 'shift'] == 2 and data_mh.loc[idx, 'hour_start'] <= 18.0 and data_mh.loc[idx, 'hour_finish'] >= 18.5:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:30:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 18:00:00')
    elif data_mh.loc[idx, 'day_cat'] == 'normal' and data_mh.loc[idx, 'shift'] == 3 and data_mh.loc[idx, 'hour_start'] <= 2.0 and data_mh.loc[idx, 'hour_finish'] >= 2.5:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 02:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 02:30:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 02:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 02:00:00')
    elif data_mh.loc[idx, 'day_cat'] == 'normal' and data_mh.loc[idx, 'shift'] == 4 and data_mh.loc[idx, 'hour_start'] <= 12.0 and data_mh.loc[idx, 'hour_finish'] >= 13.0:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 12:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 13:00:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 13:00:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 12:00:00')
    elif data_mh.loc[idx, 'day_cat'] == 'overtime' and data_mh.loc[idx, 'shift'] == 3 and data_mh.loc[idx, 'hour_start'] <= 24.0 and data_mh.loc[idx, 'hour_finish'] >= 0.5:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 00:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 00:30:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 00:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 00:00:00')
    elif data_mh.loc[idx, 'day_cat'] == 'overtime' and data_mh.loc[idx, 'shift'] == 4 and data_mh.loc[idx, 'hour_start'] <= 10.0 and data_mh.loc[idx, 'hour_finish'] >= 10.5:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 10:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 10:30:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 10:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
        data_mh.at[idx, 'date_finish'] = data_mh['date_finish'].dt.date.apply(str)[idx] + str(' 10:00:00')
    elif data_mh.loc[idx, 'shift'] == 5 and data_mh.loc[idx, 'hour_start'] <= 18.0 and data_mh.loc[idx, 'hour_finish'] >= 18.5:
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], 'Break', 'NaN', 'NaN', 'NaN', 'Break', data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:00:00'), data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:30:00'), 'NaN', 'NaN', 'NaN'])
        data.append([data_mh.loc[idx, 'name'], data_mh.loc[idx, 'year'], data_mh.loc[idx, 'month'], data_mh.loc[idx, 'week'], data_mh.loc[idx, 'day'], data_mh.loc[idx, 'day_cat'], data_mh.loc[idx, 'shift'], data_mh.loc[idx, 'task_cat'], data_mh.loc[idx, 'wo'], data_mh.loc[idx, 'machine'], data_mh['date_start'].dt.date.apply(str)[idx] + str(' 18:30:00'), data_mh.loc[idx, 'date_finish'], 'NaN', 'NaN', 'NaN'])
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
# data_mh.head(15)

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

# ---- save to excel -----
# writer = pd.ExcelWriter('data_utility.xlsx')
# data_mh.to_excel(writer, 'DATA')
# writer.save()

#%%
# ---- create plot with bokeh----
group_prod = data_mh.groupby(['name', 'month', 'task_cat']).sum()
group_prod