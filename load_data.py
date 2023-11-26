import os
import pandas as pd
import sys
import matplotlib.pyplot as plt

def read_xlsx(n):
    file_path = sys.argv[n]
    if os.path.exists(file_path):
        df = pd.read_excel(file_path, engine='openpyxl')
    else:
        print(f"File not found at: {file_path}")
        return
    return df


def map_building_id1(building_name):
    if building_name.startswith('교양분관'):
        return 0
    elif building_name.startswith('학술문화관(문화관)'):
        return 1
    elif building_name.startswith('학술문화관(도서관)'):
        return 1
    else:
        return -1  

def map_building_id2(building_name):
    if building_name.startswith('열람실'):
        return 0
    elif building_name.startswith('['):
        return 1
    else:
        return -1

def adjust_time(time_str):
    hour, minute = map(int, time_str.split(':'))
    if hour==23 and minute >=45:
        return '23:59'
    elif minute < 15: 
        minute = '00'
        
    elif minute >= 15 and minute < 45:
        minute = '30'

    elif minute >= 45:
        minute = '00'
        hour += 1

    return f"{hour:02d}:{minute}"


def make_collabo_dict(df):
    reservation_dict = {}

    for index, row in df.iterrows():
        date = row['예약일'].replace('-', '.')
        building_id = map_building_id1(row['룸(실)명칭'].split()[0])
        people = row['예약인원']

        start_time, end_time = row['예약시간'].split(' - ')
        start_time = start_time.strip()
        end_time = end_time.strip()

        reservation_list = reservation_dict.get(date, [])
        reservation_list.append((start_time, building_id, people))
        reservation_list.append((end_time, building_id, people))
        reservation_dict[date] = reservation_list

    for date, reservations in reservation_dict.items():
        reservation_dict[date] = sorted(reservations, key=lambda x: x[0])
        # print(f"Date: {date}, Reservations: {reservations}")

    return reservation_dict


def make_personal_dict(df):
    reservation_dict = {}
    for index, row in df.iterrows():
        date = row['배정일시'].split()[0] 
        start_time = adjust_time(row['시작시간'])
        end_time = adjust_time(row['마감시간'])
        building_id = map_building_id2(row['실명'].split()[0])

        reservation_list = reservation_dict.get(date, [])
        reservation_list.extend([(start_time, building_id, 1), (end_time, building_id, 1)])
        reservation_dict[date] = reservation_list

    for date, reservations in reservation_dict.items():
        reservation_dict[date] = sorted(reservations, key=lambda x: x[0])
        # print(f"Date: {date}, Reservations: {reservations}")

    return reservation_dict


def main():

    df_col = read_xlsx(1) 
    df_per = read_xlsx(2)

    reserv_dict_col = make_collabo_dict(df_col)
    reserv_dict_per = make_personal_dict(df_per)

    date = '2023.10.29'
    reservations_co = reserv_dict_col[date]
    reservations_per = reserv_dict_per[date]

    time_building_counts = {}

    for time, building, people in reservations_co:
        if (time, building) not in time_building_counts:
            time_building_counts[(time, building)] = 0
        time_building_counts[(time, building)] += people * 2

    for time, building, people in reservations_per:
        if (time, building) not in time_building_counts:
            time_building_counts[(time, building)] = 0
        time_building_counts[(time, building)] += people * 3


    sorted_times = sorted(set(time for time, _ in time_building_counts))


    buildings = [0, 1]  # 건물 ID
    for building in buildings:
        counts = [time_building_counts.get((time, building), 0) for time in sorted_times]
        plt.bar(sorted_times, counts, label=f'Building {building}')

    plt.xlabel('Time')
    plt.ylabel('Number of People')
    plt.title(f'Building Usage on {date}')
    plt.xticks(rotation=45)
    plt.legend()
    plt.show()


main()