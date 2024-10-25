from matplotlib import pyplot as plt
import matplotlib
import pytz
import datetime as dt
import os
import matplotlib.dates as mdates
from matplotlib import gridspec
import pandas as pd
import openpyxl as xl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
from openpyxl import load_workbook
import sys
import io
from matplotlib import pyplot as plt
import datetime as dt
import openpyxl as xl
import matplotlib as mp
import math 
import numpy as np
import sys
import gc
sys.setrecursionlimit(6000) 
matplotlib.use('Qt5Agg')

plt.style.use(['default'])

def make_hypnogram_to_time(hypnogram, event_start):
    hypnogram = hypnogram.tolist()
    for i in range(len(hypnogram)):
        hypnogram[i][0] = event_start + dt.timedelta(seconds=int(hypnogram[i][0]))
        hypnogram[i][1] = event_start + dt.timedelta(seconds=int(hypnogram[i][1]))
        hypnogram[i][2] = hypnogram[i][2]
    return hypnogram


def make_stage_event2(event_path, data_start):
    stage_label = {"N/A":7,"W":6,"S0":6,"N1":4,"S1":4,"N2":3,"S2":3,"N3":2,"S3":2,"R":5,"REM":5, "MT":1}
    param = ["Sleep Stage", "Position", "Time [hh:mm:ss.xxx]", "Event", "Duration[s]"]
    ind = 0
    wb = load_workbook(filename=event_path)
    ws = wb.active # 혹은 wb['Sheet1'] 같이 특정 시트 이름을 사용
    day_start = data_start.strftime("%Y-%m-%d")
    row_ind = None
    for row in ws.iter_rows(1, ws.max_row):
        if row[0].value == 'Sleep Stage':
            row_ind = row[0].row
            break    
    if row_ind is None: return
    #################### Event Start ###############################
    df = pd.read_excel(event_path, skiprows=row_ind-1)
    try:
        now = df[param[2]].tolist()[0]
        event_start_s = day_start +" "+ str(now.hour)+":"+str(now.minute)+":"+str(now.second)
    except:
        param[2] = "Time [hh:mm:ss]"
        now = df[param[2]].tolist()[0]
        event_start_s = day_start +" "+ str(now.hour)+":"+str(now.minute)+":"+str(now.second)

    format= "%Y-%m-%d %H:%M:%S"

    event_start = dt.datetime.strptime(event_start_s, format).replace(tzinfo=pytz.utc)
    if event_start.hour <12 : event_start += dt.timedelta(hours=12)
    ##################### Events calibration #########################
    times_s = df[param[2]].tolist()
    times = []
    for time_s in times_s:
        time_dt = dt.datetime.strptime(day_start +" "+ str(time_s.hour)+":"+str(time_s.minute)+":"+str(time_s.second), format).replace(tzinfo=pytz.utc)
        if time_dt.hour < 9 and time_dt < event_start:
            time_dt += dt.timedelta(hours=24)
        if time_dt < event_start:
            time_dt += dt.timedelta(hours = 12)
        
        times.append((time_dt-event_start).seconds)
    ##################### Hypnogram #################################
    stages = df[param[0]].tolist()
    duration = [round(dur) for dur in df[param[4]].tolist()]
    events = []
    hypnogram = np.empty((0,3), int)
    
    for i in range(len(stages)):
        event= 7
        try:
            for key in stage_label.keys():
                if key in stages[i].upper(): event = stage_label[key]
        except:
            pass
        events.append(event)

    for i in range(len(stages)):
        if i == 0 or events[i] != events[i-1]:
            ev_start = times[i]
        elif i == len(stages)-1:
            ev_end = times[i] + duration[i]
            hypnogram = np.append(hypnogram, np.array([[ev_start, ev_end, events[i]]]), axis = 0)
        elif events[i] != events[i+1]:
            ev_end = times[i+1]
            hypnogram = np.append(hypnogram, np.array([[ev_start, ev_end, events[i]]]), axis = 0)
    
    ##################### REM #################################
    REM_events = np.empty((0,2), int)
    
    for hyp in hypnogram:
        if hyp[2] == 5:
            REM_events = np.append(REM_events, np.array([[hyp[0],hyp[1]]]), axis = 0)
    ##################### start param #################################         
    start_index = (event_start - data_start).seconds * 200
    if event_start < data_start : start_index = -(data_start - event_start).seconds * 200
    hypnogram = make_hypnogram_to_time(hypnogram, event_start)
    return [REM_events, start_index, event_start,hypnogram,format]
    # sleep_event : [start_time, end_time]
    
    
    
def make_stage_event(event_path, data_start):
    stage_label = {"N/A":7,"W":6,"S0":6,"N1":4,"S1":4,"N2":3,"S2":3,"N3":2,"S3":2,"R":5,"REM":5, "MT":1}
    param = ["Sleep Stage", "Position", "Time [hh:mm:ss.xxx]", "Event", "Duration[s]"]
    ind = 1
    wb = load_workbook(filename=event_path)
    ws = wb.active # 혹은 wb['Sheet1'] 같이 특정 시트 이름을 사용
    day_start = data_start.strftime("%Y-%m-%d")
    row_ind = None
    for row in ws.iter_rows(1, ws.max_row):
        if row[0].value == 'Sleep Stage':
            row_ind = row[0].row
            break    
    if row_ind is None: return
    #################### Event Start ###############################
    df = pd.read_excel(event_path, skiprows=row_ind-1)
    try:
        if df[param[2]].tolist()[0].split(" ")[0] in ":": ind = 0
        event_start_s = day_start +" "+ df[param[2]].tolist()[0].split(" ")[ind].split(".")[0]
    except:
        param[2] = "Time [hh:mm:ss]"
        if df[param[2]].tolist()[0].split(" ")[0] in ":": ind = 0
        event_start_s = day_start +" "+ df[param[2]].tolist()[0].split(" ")[ind].split(".")[0]

    format= "%Y-%m-%d %H:%M:%S"

    event_start = dt.datetime.strptime(event_start_s, format).replace(tzinfo=pytz.utc)
    if event_start.hour <12 : event_start += dt.timedelta(hours=12)
    ##################### Events calibration #########################
    times_s = df[param[2]].tolist()
    times = []
    for time_s in times_s:
        time_dt = dt.datetime.strptime(day_start +" "+ time_s.split(" ")[ind].split(".")[0], format).replace(tzinfo=pytz.utc)
        if time_dt.hour < 9 and time_dt < event_start:
            time_dt += dt.timedelta(hours=24)
        if time_dt < event_start:
            time_dt += dt.timedelta(hours = 12)
        
        times.append((time_dt-event_start).seconds)
    ##################### Hypnogram #################################
    stages = df[param[0]].tolist()
    duration = [round(dur) for dur in df[param[4]].tolist()]
    events = []
    hypnogram = np.empty((0,3), int)
    
    for i in range(len(stages)):
        event= 7
        try:
            for key in stage_label.keys():
                if key in stages[i].upper(): event = stage_label[key]
        except:
            pass
        events.append(event)

    for i in range(len(stages)):
        if i == 0 or events[i] != events[i-1]:
            ev_start = times[i]
        elif i == len(stages)-1:
            ev_end = times[i] + duration[i]
            hypnogram = np.append(hypnogram, np.array([[ev_start, ev_end, events[i]]]), axis = 0)
        elif events[i] != events[i+1]:
            ev_end = times[i+1]
            hypnogram = np.append(hypnogram, np.array([[ev_start, ev_end, events[i]]]), axis = 0)
    
    ##################### REM #################################
    REM_events = np.empty((0,2), int)
    
    for hyp in hypnogram:
        if hyp[2] == 5:
            REM_events = np.append(REM_events, np.array([[hyp[0],hyp[1]]]), axis = 0)
    ##################### start param #################################         
    start_index = (event_start - data_start).seconds * 200
    if event_start < data_start : start_index = -(data_start - event_start).seconds * 200
    hypnogram = make_hypnogram_to_time(hypnogram, event_start)
    return [REM_events, start_index, event_start,hypnogram,format]
    # sleep_event : [start_time, end_time]
    
def make_artifact2(event_path, event_start,format, REM):
    AHI_list = ['apnea','hypopnea']
    artifact = ['artifact_chin','artifact_rarm','artifact_larm','artifact_rleg','artifact_lleg']
    artifact2 = ['artifact_chin','artifact_rfds','artifact_lfds','artifact_rta','artifact_lta']
    lists = [[],[],[],[],[]]
    AHIs = []

    param = ["Sleep Stage", "Position", "Time [hh:mm:ss.xxx]", "Event", "Duration[s]"]
    wb = load_workbook(filename=event_path)
    ws = wb.active # 혹은 wb['Sheet1'] 같이 특정 시트 이름을 사용
    day_start = event_start.strftime("%Y-%m-%d")
    row_ind = None
    for row in ws.iter_rows(1, ws.max_row):
        if row[0].value == 'Sleep Stage':
            row_ind = row[0].row
            break    
    if row_ind is None: return

    df = pd.read_excel(event_path, skiprows=row_ind-1)

    ##################### Events calibration #########################
    try:
        times_s = df[param[2]].tolist()
    except:
        param[2] ="Time [hh:mm:ss]"
        times_s = df[param[2]].tolist()
    ind = 1
    times = []
    for time_s in times_s:
        time_dt = dt.datetime.strptime(day_start +" "+ str(time_s.hour)+":"+str(time_s.minute)+":"+str(time_s.second), format).replace(tzinfo=pytz.utc)
        if time_dt.hour < 9 and time_dt < event_start:
            time_dt += dt.timedelta(hours=24)
        if time_dt < event_start:
            time_dt += dt.timedelta(hours = 12)
        
        times.append((time_dt-event_start).seconds)
    ##################### Hypnogram #################################
    events = df[param[3]].tolist()
    duration = [round(dur) for dur in df[param[4]].tolist()]
   
    for i in range(5):
        for j in range(len(events)):
            if duration[j] > 300: continue
            if 'arousal' in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if 'rera' in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if artifact[i] in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if artifact2[i] in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            
            if i != 0: continue
            if 'snore' in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if events[j].lower() in AHI_list: AHIs.append([times[j],times[j]+ duration[j]])

        
        AHI = calc_AHI(AHIs, REM)
    if AHI > 15:
        for i in range(5):
            for j in range(len(events)):
                if 'apnea' in events[j]: lists[i].append([times[j],times[j]+ duration[j]])
                if 'hypopnea' in events[j]: lists[i].append([times[j],times[j]+ duration[j]])      
                
            lists[i]= np.array(lists[i])  
    return [lists, AHI]
    # sleep_event : [start_time, end_time]
        
def make_artifact(event_path, event_start,format, REM):
    artifact = ['artifact_chin','artifact_rarm','artifact_larm','artifact_rleg','artifact_lleg']
    artifact2 = ['artifact_chin','artifact_rfds','artifact_lfds','artifact_rta','artifact_lta']
    lists = [[],[],[],[],[]]
    AHIs = []

    param = ["Sleep Stage", "Position", "Time [hh:mm:ss.xxx]", "Event", "Duration[s]"]
    wb = load_workbook(filename=event_path)
    ws = wb.active # 혹은 wb['Sheet1'] 같이 특정 시트 이름을 사용
    day_start = event_start.strftime("%Y-%m-%d")
    row_ind = None
    for row in ws.iter_rows(1, ws.max_row):
        if row[0].value == 'Sleep Stage':
            row_ind = row[0].row
            break    
    if row_ind is None: return

    df = pd.read_excel(event_path, skiprows=row_ind-1)

    ##################### Events calibration #########################
    try:
        times_s = df[param[2]].tolist()
    except:
        param[2] ="Time [hh:mm:ss]"
        times_s = df[param[2]].tolist()
    ind = 1
    if ":" in times_s[0].split(" ")[0]: ind = 0
    times = []
    for time_s in times_s:
        time_dt = dt.datetime.strptime(day_start +" "+ time_s.split(" ")[ind].split(".")[0], format).replace(tzinfo=pytz.utc)
        if time_dt.hour < 9 and time_dt < event_start:
            time_dt += dt.timedelta(hours=24)
        if time_dt < event_start:
            time_dt += dt.timedelta(hours = 12)
        
        times.append((time_dt-event_start).seconds)
    ##################### Hypnogram #################################
    events = df[param[3]].tolist()
    duration = [round(dur) for dur in df[param[4]].tolist()]
   


    for i in range(5):
        for j in range(len(events)):
            if duration[j] > 300: continue
            if 'arousal' in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if 'rera' in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if artifact[i] in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if artifact2[i] in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            
            if i != 0: continue
            if 'snore' in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if 'apnea' in events[j].lower() : AHIs.append([times[j],times[j]+ duration[j]])
            if 'hypopnea' in events[j].lower() : AHIs.append([times[j],times[j]+ duration[j]])

        
    AHI = calc_AHI(AHIs, REM)
    if AHI > 15:
        for i in range(5):
            for j in range(len(events)):
                if 'apnea' in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
                if 'hypopnea' in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])   
        
            lists[i]= np.array(lists[i])     
    return [lists, AHI]
    # sleep_event : [start_time, end_time]
    
import re
def clean_string(s):
    return re.sub(r'[^0-9.:]', '', s)

def make_stage_event3(event_path, data_start):
    stage_label = {"N/A":7,"W":6,"S0":6,"N1":4,"S1":4,"N2":3,"S2":3,"N3":2,"S3":2,"R":5,"REM":5, "MT":1}
    param = ["Sleep Stage", "Position", "Time [hh:mm:ss.xxx]", "Event", "Duration[s]"]
    ind = 0
    wb = load_workbook(filename=event_path)
    ws = wb.active # 혹은 wb['Sheet1'] 같이 특정 시트 이름을 사용
    day_start = data_start.strftime("%Y-%m-%d")
    row_ind = None
    for row in ws.iter_rows(1, ws.max_row):
        if row[0].value == 'Sleep Stage':
            row_ind = row[0].row
            break    
    if row_ind is None: return
    #################### Event Start ###############################
    df = pd.read_excel(event_path, skiprows=row_ind-1)
    df[param[2]] = df[param[2]].apply(lambda x: x.strftime('%H:%M:%S.%f') if isinstance(x, pd.Timestamp) else x)
    df[param[2]] = df[param[2]].apply(clean_string)
    df[param[2]] = pd.to_datetime(df[param[2]], format='%H:%M:%S.%f').dt.time
    try:
        now = df[param[2]].tolist()[0]
        try:
            event_start_s = day_start +" "+ str(now.hour)+":"+str(now.minute)+":"+str(now.second)
        except:
            event_start_s = day_start +" " + now
    except:
        param[2] = "Time [hh:mm:ss]"
        try:
            event_start_s = day_start +" "+ str(now.hour)+":"+str(now.minute)+":"+str(now.second)
        except:
            event_start_s = day_start +" " + now

    format= "%Y-%m-%d %H:%M:%S"

    event_start = dt.datetime.strptime(event_start_s, format).replace(tzinfo=pytz.utc)
    if 12 >= event_start.hour >=6 : event_start += dt.timedelta(hours=12)
    ##################### Events calibration #########################
    times_s = df[param[2]].tolist()
    times = []
    for time_s in times_s:
        time_dt = dt.datetime.strptime(day_start +" "+ str(time_s.hour)+":"+str(time_s.minute)+":"+str(time_s.second), format).replace(tzinfo=pytz.utc)
        if time_dt.hour < 9 and time_dt < event_start:
            time_dt += dt.timedelta(hours=24)
        if time_dt < event_start:
            time_dt += dt.timedelta(hours = 12)
        
        times.append((time_dt-event_start).seconds)
    ##################### Hypnogram #################################
    stages = df[param[0]].tolist()
    duration = [round(dur) for dur in df[param[4]].tolist()]
    events = []
    hypnogram = np.empty((0,3), int)
    
    for i in range(len(stages)):
        event= 7
        try:
            for key in stage_label.keys():
                if key in stages[i].upper(): event = stage_label[key]
        except:
            pass
        events.append(event)

    for i in range(len(stages)):
        if i == 0 or events[i] != events[i-1]:
            ev_start = times[i]
        elif i == len(stages)-1:
            ev_end = times[i] + duration[i]
            hypnogram = np.append(hypnogram, np.array([[ev_start, ev_end, events[i]]]), axis = 0)
        elif events[i] != events[i+1]:
            ev_end = times[i+1]
            hypnogram = np.append(hypnogram, np.array([[ev_start, ev_end, events[i]]]), axis = 0)
    
    ##################### REM #################################
    REM_events = np.empty((0,2), int)
    
    for hyp in hypnogram:
        if hyp[2] == 5:
            REM_events = np.append(REM_events, np.array([[hyp[0],hyp[1]]]), axis = 0)
    ##################### start param #################################         
    start_index = (event_start - data_start).seconds * 200
    if event_start < data_start : start_index = -(data_start - event_start).seconds * 200
    hypnogram = make_hypnogram_to_time(hypnogram, event_start)
    return [REM_events, start_index, event_start,hypnogram,format]
    # sleep_event : [start_time, end_time]
    
 
    # sleep_event : [start_time, end_time]
def make_artifact3(event_path, event_start,format, REM):
    AHI_list = ['apnea','hypopnea']
    artifact = ['artifact_chin','artifact_rarm','artifact_larm','artifact_rleg','artifact_lleg']
    artifact2 = ['artifact_chin','artifact_rfds','artifact_lfds','artifact_rta','artifact_lta']
    lists = [[],[],[],[],[]]
    AHIs = []

    param = ["Sleep Stage", "Position", "Time [hh:mm:ss.xxx]", "Event", "Duration[s]"]
    wb = load_workbook(filename=event_path)
    ws = wb.active # 혹은 wb['Sheet1'] 같이 특정 시트 이름을 사용
    day_start = event_start.strftime("%Y-%m-%d")
    row_ind = None
    for row in ws.iter_rows(1, ws.max_row):
        if row[0].value == 'Sleep Stage':
            row_ind = row[0].row
            break    
    if row_ind is None: return

    df = pd.read_excel(event_path, skiprows=row_ind-1)
    df[param[2]] = df[param[2]].apply(clean_string)
    df[param[2]] = pd.to_datetime(df[param[2]], format='%H:%M:%S.%f').dt.time


    ##################### Events calibration #########################
    try:
        times_s = df[param[2]].tolist()
    except:
        param[2] ="Time [hh:mm:ss]"
        times_s = df[param[2]].tolist()
    ind = 1
    times = []
    for time_s in times_s:
        time_dt = dt.datetime.strptime(day_start +" "+ str(time_s.hour)+":"+str(time_s.minute)+":"+str(time_s.second), format).replace(tzinfo=pytz.utc)
        if time_dt.hour < 9 and time_dt < event_start:
            time_dt += dt.timedelta(hours=24)
        if time_dt < event_start:
            time_dt += dt.timedelta(hours = 12)
        
        times.append((time_dt-event_start).seconds)
    ##################### Hypnogram #################################
    events = df[param[3]].tolist()
    duration = [round(dur) for dur in df[param[4]].tolist()]
   
    for i in range(5):
        for j in range(len(events)):
            if duration[j] > 300: continue
            if 'arousal' in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if 'rera' in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if artifact[i] in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if artifact2[i] in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            
            if i != 0: continue
            if 'snore' in events[j].lower(): lists[i].append([times[j],times[j]+ duration[j]])
            if 'apnea' in events[j].lower() : AHIs.append([times[j],times[j]+ duration[j]])
            if 'hypopnea' in events[j].lower() : AHIs.append([times[j],times[j]+ duration[j]])

        
        AHI = calc_AHI(AHIs, REM)
    if AHI > 15:
        for i in range(5):
            for j in range(len(events)):
                if 'apnea' in events[j]: lists[i].append([times[j],times[j]+ duration[j]])
                if 'hypopnea' in events[j]: lists[i].append([times[j],times[j]+ duration[j]])      
                
            lists[i]= np.array(lists[i])  
    return [lists, AHI]
    # sleep_event : [start_time, end_time]

    
def calc_AHI( AHIs, REMs):
    total_AHI = total_REM = 0
    for rem in REMs:
        total_REM += rem[1]-rem[0]      
        for ahi in AHIs:
            if rem[0] <= ahi[0] < ahi[1]<=rem[1]:
                total_AHI += 1
            elif rem[0] <= ahi[0] < rem[1] <= ahi[1]:
                total_AHI += 1
            elif ahi[0] <=rem[0] < rem[1] <= ahi[1]:
                total_AHI += 1
            elif ahi[0]<= rem[0]<ahi[1]<rem[1]:
                total_AHI += 1
    
    AHI = round(total_AHI /total_REM *60*60,3)
    
    return AHI
    
        


def make_REM_period(events):
    REM = [np.empty((0,2),int) for _ in events]
    
    for i in range(len(events)) :
        REM[i] = np.append(REM[i], np.array([events[i][0].tolist()]), axis = 0)
        for j in range(1, len(events[i])):
            if REM[i][-1][1] == events[i][j][0]:
                REM[i][-1][1] = events[i][j][1]
            else :
                REM[i] = np.append(REM[i], np.array([events[i][j].tolist()]), axis = 0)

    
    return REM

def fit_to_channel(channel,REM, l):
    fit_REM = REM.copy()
    for i in range(len(channel)):
        if channel[i] == False: fit_REM[i] = [1 for _ in range(l)]
        
    return fit_REM
    

def filter_Seperate(data, start_index, channels):
    
    pickss = [0,0,0,0,0]
    ch_names = []
    for i in range(len(channels)):
        if ("chin" in channels[i].lower()) or ("mentalis" in channels[i].lower()) or ("lower.left" in channels[i].lower()):
            pickss[0] = channels[i]
            ch_names.append(channels[i])
        elif ("leg" in channels[i].lower()) or ('tibialis' in channels[i].lower()):
            if 'r' in channels[i].lower() :
                pickss[1] = channels[i]
                ch_names.append(channels[i])
            elif 'l' in channels[i].lower():
                pickss[2] =channels[i]
                ch_names.append(channels[i])
        elif ("arm" in channels[i].lower()) or ('fds' in channels[i].lower()):
            if 'l' in channels[i].lower(): pickss[4] =channels[i];ch_names.append(channels[i])
            else: pickss[3] = channels[i];ch_names.append(channels[i])
    
    picks = [item for item in pickss if item != 0]
    datas = data.copy().pick_channels(picks, ordered = True)
    noise_freq = 60
    datas.notch_filter(noise_freq)
    data_filt = datas.copy().filter(l_freq=10, h_freq = 100)
    data_res = data_filt.resample(200)
    f_EMG_data = data_res.get_data()   #make filtered EMG_data
    RTA_EMG = []
    LTA_EMG = []
    Rarm_EMG = []
    Larm_EMG = []
    chin_EMG = []
    EMG = [chin_EMG, RTA_EMG, LTA_EMG, Rarm_EMG, Larm_EMG  ]

    channel = [False, False, False, False, False]
    ch = ["Chin","Rt_TA","Lt_TA","Rt_FDS","Lt_FDS"]

    if start_index <0:
        f_EMG_data = np.pad(f_EMG_data, ((0, 0), (-start_index, 0)), mode='constant', constant_values=0)
        start_index = 0
    for i in range(len(picks)):
        if ("chin" in picks[i].lower()) or ("mentalis" in picks[i].lower()) or ("lower.left" in picks[i].lower()):
            if (channel[0] == False):
                EMG[0] = f_EMG_data[i][start_index:]
                channel[0] = True
                
        elif ("leg" in picks[i].lower()) or ('tibialis' in picks[i].lower()):
            if ('r' in picks[i].lower() ) :
                EMG[1] = f_EMG_data[i][start_index:]
                channel[1] = True
            else :
                EMG[2] = f_EMG_data[i][start_index:]
                channel[2] = True
                
        elif ("arm" in picks[i].lower()) or ('fds' in picks[i].lower()):
            if ('l' in picks[i].lower() ):
                EMG[4] = f_EMG_data[i][start_index:]
                channel[4] = True
            else:
                EMG[3] = f_EMG_data[i][start_index:]
                channel[3] = True           


    EMG = fit_to_channel(channel, EMG, len(f_EMG_data[0][start_index:])) # channel * timepoints
    # data_res.plot(start = int(start_index/256),time_format = 'clock')
    # input()
    trg = [True for _ in range(5)]
    if channel[1] == False or channel[2] == False:
        trg[1] = False; trg[4] = False
    if channel[3] == False or channel[4] == False:
        trg[2] = False; trg[3] = False;trg[4] = False

    return [EMG, ch, trg]
def make_REM(REMs, artifacts):
    epochs = [np.empty((0,2),int) for _ in artifacts] 

    
    for i in range(len(artifacts)):
        for R in REMs:
            t_end = start = R[0]
            end = R[1]
            
            while t_end < end :
                if (end - t_end) >= 30:
                    t_end+=30
                    epochs[i] = np.append(epochs[i], np.array([[start,t_end]]),axis =0)
                    start+=30
                else : break

    return epochs

def make_REM_epochs(REMs, artifacts):
    epochs = [np.empty((0,2),int) for _ in artifacts]
    mini = [np.empty((0,2),int) for _ in artifacts] 
    base_trgs= [[] for _ in artifacts] 
    trgs = []
    
    for i in range(len(artifacts)):
        for R in REMs:
            t_end = start = R[0]
            end = R[1]
            
            while t_end < end :
                if (end - t_end) >= 30:
                    t_end+=30
                    trg = 1
                    for arfs in artifacts[i]:
                        if (start <= arfs[0] <= t_end): trg=0
                        if (start <= arfs[1] <= t_end): trg=0
                        if (arfs[0] <= start <= arfs[1]): trg=0
                        if (arfs[0] <= t_end <= arfs[1]): trg=0
                    
                    base_trgs[i].append(trg)
                    epochs[i] = np.append(epochs[i], np.array([[start,t_end]]),axis =0)
                    start+=30
                else : break
                
    for i in range(len(artifacts)):
        ok_base = False
        oks =0
        for trg in base_trgs[i]:
            if trg == 1:
                oks+=1
            else: 
                oks = 0
            if oks == 3:
                ok_base = True
                break
        
        if ok_base == True : 
            trgs.append(0)
            continue
        epochs[i] = np.empty((0,2),int)
        base_trgs[i] = []
        trgs.append(1)

        for R in REMs:
            t_end = start = R[0]
            end = R[1]
            
            while t_end < end :
                if (end - t_end) >= 30:
                    t_end+=30
                    trg = 1
                    for arfs in artifacts[i]:
                        if (start <= arfs[0] <= t_end): trg=0
                        if (start <= arfs[1] <= t_end): trg=0
                        if (arfs[0] <= start <= arfs[1]): trg=0
                        if (arfs[0] <= t_end <= arfs[1]): trg=0
                    
                    base_trgs[i].append(0)
                    epochs[i] = np.append(epochs[i], np.array([[start,t_end]]),axis =0)
                    start+=30
                else : break   
    a= 2 

    return [epochs, base_trgs, trgs]


def RMS(E):
    # E2 = np.power(E,2)
    # window = np.ones(len(E))/float(len(E))
    rms = np.sqrt(np.mean(np.square(E)))

    # return rms
    return rms

def make_rms(EMG, epochs , f_s):
    epoch_rms = []
    for i in range(len(EMG)): 
        epoch_rms.append(np.zeros(len(epochs[i])))
    epoch_rms = np.array(epoch_rms, dtype=object)
    
    i=0

    for E in EMG:  #make rms
        j=0
        
        for epoch in epochs[i]:
            epoch_start = epoch[0] * f_s
            epoch_end   = epoch[1] * f_s

            rms = RMS(E[epoch_start : epoch_end])
            if rms < 0.05e-06:
                rms = 11
            # if rms > 5e-06:
            #     rms = 11
            epoch_rms[i][j] = rms
            
            j+=1
        

        i+=1
    return epoch_rms


def make_baseline(epoch_rms,REMs, epochs, art_epochs, trgs):
    baselines = []
    res_baselines = []
    for i in range(len(epochs)):
        baselines.append(np.ones(len(epochs[i])))
        res_baselines.append(np.ones(len(epochs[i])))
    baselines = np.array(baselines, dtype= object)
    res_baselines = np.array(baselines, dtype= object)
    
    period= 90

    for i in range(len(epochs)):
        length =  epochs[i][0][1] - epochs[i][0][0] 
        num = period // length
        eps = []
        bases = []
        for j in range(len(epochs[i])):
            if art_epochs[i][j] ==1: #1이 없는거
                eps.append(j)

        ind = 0
        for j in range(1,len(eps)):
            # if eps[j] - eps[j-1] != 1 or eps[j] - ind == num:
            if epochs[i][eps[j]][0] - epochs[i][eps[j-1]][1] > 1:
                if eps[j] - ind >= num:
                    base = np.min(epoch_rms[i][ind:eps[j]])
                    for k in range(ind,eps[j]):
                        baselines[i][k] = base
                ind = eps[j]
                
        if ind == 0:
            for k in range(len(baselines[i])):
                if len(eps) == 0: baselines[i][k] = 11; continue
                baselines[i][k] = np.min(epoch_rms[i])
    l = 0
    for baseline in baselines:
        for i in range(len(baseline)):
            if baseline[i] == 1:
                L_index = len(epochs)-1
                R_index = 0
                for j in range(len(baseline)):
                    if   j<i and baseline[j] != 1 : L_index = j
                    elif j>i and baseline[j] != 1 : R_index = j; break
                diff_L = epochs[l][i][0] - epochs[l][L_index][1]
                diff_R = epochs[l][R_index][0] - epochs[l][i][1]
                
                if diff_R > diff_L:
                    res_baselines[l][i] = baseline[R_index]
                elif diff_L > diff_R:
                    res_baselines[l][i] = baseline[L_index]
                elif diff_L == diff_R:
                    res_baselines[l][i] = (baseline[R_index] + baseline[L_index])/2
                # if diff_L <= period and diff_R <= period:
                #     baseline[i] = (baseline[R_index] + baseline[L_index])/2
                # elif diff_L <= period and diff_R > period:
                #     baseline[i] = baseline[L_index]
                # elif diff_R <= period and diff_L > period :
                #     baseline[i] = baseline[R_index]
            else:
                res_baselines[l][i] = baseline[i]
        l+=1
        
    
    for i in range(len(res_baselines)):
        for j in range(len(baselines)):
            if epoch_rms[i][j] == 11:
                baselines[j] = 11
        
    return res_baselines


def make_activity(baselines, artifacts , EMGs, f_s ,epochs ) :
    bouts = int(0.03 * f_s)
    activitys  = [np.empty((0,3),int) for _ in range(len(baselines))]
    activities  = [np.empty((0,3),int) for _ in range(len(baselines))]
    res_activities  = [np.empty((0,3),int) for _ in range(len(baselines))]
    
    for i in range(len(baselines)):
        for j in range(len(baselines[i])):
            baseline2 = baselines[i][j]*2
            start = epochs[i][j][0] * f_s
            end = epochs[i][j][1] * f_s
            t_end = start + bouts

            up_duration = 0
            down_duration = 0
            odd_trg= False
            while t_end <= end :
                rms = RMS(EMGs[i][start : t_end])
                if up_duration <= 0.1 * f_s:
                    if rms >= baseline2:
                        if len(activitys[i]) != 0 and start - activitys[i][-1][1] < 0.25*f_s:
                            activity_start = int(start - 0.105*f_s)
                            up_duration = int(0.105*f_s)
                            odd_trg = True
                        if up_duration == 0:
                            activity_start = start
                            up_duration += int(0.015 * f_s)
                        if activity_start != 0:
                            down_duration = 0
                            up_duration += int(0.015 * f_s)
                    else:
                        up_duration =0
                else:
                    if rms >= baseline2:
                        up_duration += int(0.015 * f_s)
                        down_duration = 0
                    else :
                        if down_duration == 0:
                            down_duration += int(0.015 * f_s)
                        else:
                            down_duration += int(0.015 * f_s)
                        if down_duration >= 0.25*f_s :
                            if odd_trg == False:
                                RMS_now = RMS(EMGs[i][activity_start: t_end-down_duration])
                                RMS_previous = RMS(EMGs[i][round(activity_start-0.25*f_s): activity_start])
                                if RMS_now > RMS_previous*2 :
                                    activitys[i] = np.append(activitys[i],np.array([[activity_start,t_end-down_duration,1]]),axis =0 )
                            else:
                                activitys[i] = np.append(activitys[i],np.array([[activity_start,t_end-down_duration,1]]),axis =0 )
                            up_duration = 0 
                            odd_trg = False       
                start += int(0.015 * f_s)
                t_end += int(0.015 * f_s)
            if up_duration > 0.1 * f_s :
                activitys[i] = np.append(activitys[i],np.array([[activity_start,t_end-down_duration,1]]),axis =0 )
        if len(activitys[i]) > 0 :
            activities[i] = np.append(activities[i], np.array([activitys[i][0]]), axis=0)
            for k in range(1,len(activitys[i])):
                if activitys[i][k][0] - activities[i][-1][1] <= 0.25*f_s:
                    activities[i][-1][1] = activitys[i][k][1]
                else:
                    activities[i] = np.append(activities[i], np.array([activitys[i][k]]), axis=0)
        
        for j in range(len(activities[i])):
            act = activities[i][j]
            start = act[0] // f_s
            end = act [1] // f_s
            trg = 1
            # if len(activitys[i]) != 0 :
            #     RMS_now = RMS(EMGs[i][activities[i][j][0]: activities[i][j][1]])
            #     RMS_previous = RMS(EMGs[i][activities[i][j][0] - int(0.25*f_s): activities[i][j][0]])
            #     if RMS_now > RMS_previous*2 :
            #         activitys[i] = np.append(activitys[i],np.array([[activity_start,t_end-down_duration,1]]),axis =0 )
            for emg in abs(EMGs[i][act[0]:act[1]]):
                if emg > 250e-6: trg = 0
            for art in artifacts[i]:
                if start<= art[1] <= end : trg = 0
                if start<= art[0] <= end : trg = 0 
                if art[0]<= start <= art[1]: trg = 0
                if art[0]<= end <= art[1]: trg = 0
            if trg ==1:
                res_activities[i] = np.append(res_activities[i], np.array([act]), axis=0)
        a=1
    return res_activities
                
def make_event(activitys, epochs, f_s, artifacts):
    events30  = [np.empty((0,3),int) for _ in range(len(activitys))]
    events3  = [np.empty((0,3),int) for _ in range(len(activitys))]
    tonic_duration = 0
    i_cut = tonic_cut = 5 * f_s

    for i in range(len(activitys)):
        for j in range(len(epochs[i])):
            epoch = epochs[i][j]
            e_start = epoch[0] * f_s
            e_end   = epoch[1] * f_s
            
            duration = 0
            for j in range(len(activitys[i])):
                a_start = activitys[i][j][0]
                a_end   = activitys[i][j][1] 
                if a_end - a_start < 5*f_s : continue
                if  a_start <= e_start < a_end <= e_end :
                    duration += a_end - e_start
                if e_start < a_start < a_end <= e_end  :
                    duration += a_end - a_start  
                if e_start <= a_start < e_end <= a_end :
                    duration += e_end - a_start           
                if a_start <= e_start < e_end <= a_end :
                    duration += e_end - e_start

            art_trg = 0    
            for t in range(len(artifacts[i])):
                art = artifacts[i][t]
                if art[0]*200 <= e_start <= art[1]*200: art_trg = 1; break
                if art[0]*200 <= e_end <=art[1]*200: art_trg = 1; break  
                if  e_start<= art[0]*200 <=e_end: art_trg = 1; break  
            if  duration >= 15*f_s :
                events30[i]  = np.append(events30[i], np.array([[e_start,e_end,0]]),axis =0 )
            elif art_trg == 1:
                events30[i]  = np.append(events30[i], np.array([[e_start,e_end,11]]),axis =0 )
            else:
                events30[i]  = np.append(events30[i], np.array([[e_start,e_end,10]]),axis =0 )
                
                
    
    for i in range(len(activitys)):
        for j in range(len(epochs[i])):
            epoch = epochs[i][j]
            e_start = epoch[0] * f_s
            e_end   = epoch[1] * f_s

            e_t_start = e_start
            e_t_end = epoch[0]*f_s+3* f_s

            if events30[i][j][2] ==0: 
                for _ in range(10):
                    events3[i]  = np.append(events3[i], np.array([[e_t_start,e_t_end,0]]),axis =0 )
                    e_t_start += 3*f_s
                    e_t_end += 3*f_s
                continue
            
            while e_t_end <= e_end:
                tonic_duration = 0
                duration = 0
                for t in range(len(activitys[i])):
                    a_start = activitys[i][t][0]
                    a_end   = activitys[i][t][1] 
                    if  a_start <= e_t_start < a_end <= e_t_end :
                        duration += a_end - e_t_start
                        if a_end - a_start >= i_cut: tonic_duration += a_end - e_t_start
                    if e_t_start < a_start < a_end <= e_t_end  :
                        duration += a_end - a_start  
                        if a_end - a_start >= i_cut: tonic_duration += a_end - a_start             
                    if e_t_start <= a_start < e_t_end <= a_end :
                        duration += e_t_end - a_start           
                        if a_end - a_start >= i_cut: tonic_duration += e_t_end - a_start
                    if a_start <= e_t_start < e_t_end <= a_end :
                        duration += e_t_end - e_t_start
                        if a_end - a_start >= i_cut: tonic_duration += e_t_end - e_t_start 
                art_trg = 0    
                for t in range(len(artifacts[i])):
                    art = artifacts[i][t]
                    if art[0]*200 <= e_t_start <= art[1]*200: art_trg = 1; break
                    if art[0]*200 <= e_t_end <=art[1]*200: art_trg = 1; break 
                    if  e_start<= art[0]*200 <=e_end: art_trg = 1; break  
                if tonic_duration > 0:
                    events3[i]  = np.append(events3[i], np.array([[e_t_start,e_t_end,2]]),axis =0 )
                elif duration > 0 :
                    events3[i]  = np.append(events3[i], np.array([[e_t_start,e_t_end,1]]),axis =0 ) 
                elif art_trg == 1:
                    events3[i]  = np.append(events3[i], np.array([[e_t_start,e_t_end,11]]),axis =0 )
                else:
                    events3[i]  = np.append(events3[i], np.array([[e_t_start,e_t_end,10]]),axis =0 ) 

                e_t_start += 3*f_s
                e_t_end += 3*f_s

            count = 0
            all_count = 0
            for l in range(10):
                if events3[i][-l-1][2] == 1:
                    count += 1
                if events3[i][-l-1][2] not in [10,11]:
                    all_count +=1
            if count >= 5:
                events30[i][j][2] = 1
            elif all_count >=5:
                events30[i][j][2] = 2
        A= 1
# 1 tonic 0 phasic 2 any 10 none 11 ART
            
    return events3, events30


    
def make_activity2( EMG, f_s ,event):
    bouts = int(0.02 * f_s)
    activitys  = [np.empty((0,3),int) for i in range(len(EMG))]
    for j in range(len(EMG)):
        if j == 0: lm = 0
        if j == 1: lm = 1
        if j == 2: lm = 1
        if j == 3: lm = 2
        elif j == 4: lm = 2
        
        for i in range(len(event[lm])):
            if event[lm][i][2] == 0:
                ss = start = event[lm][i][0]
                t_end = start + bouts
                end   = start + 30*f_s
                up_duration = 0
                down_duration = 0
                baseline4 = float(RMS(EMG[j][ss : end]))*2
                while t_end <= end :
                    rms = RMS(EMG[j][start : t_end])
                    if up_duration < 5*bouts:
                        if rms >= baseline4:
                            if up_duration == 0 :
                                activity_start = start
                                up_duration +=int(0.01 * f_s)
                            if activity_start != 0:
                                down_duration = 0
                                up_duration += int(0.01 * f_s)
                        else:
                            up_duration =0
                    else:
                        if rms >= baseline4:
                            up_duration += int(0.01 * f_s)
                            down_duration = 0
                        else :                      
                            down_duration += int(0.01 * f_s)
                            if down_duration >= 0.25*f_s :
                                activitys[j] = np.append(activitys[j],np.array([[activity_start,t_end-down_duration,1]]),axis =0 )
                                up_duration = 0        
                                
                    start += int(0.01 * f_s)
                    t_end +=int(0.01 * f_s)
                if up_duration >=5*bouts :
                    activitys[j] = np.append(activitys[j],np.array([[activity_start,t_end-down_duration,1]]),axis =0 )
                i += 10

                
    return activitys    


def make_event2(activitys, f_s,event): 
    for i in range(len(event)): 
        for e in event[i]:
            if e[2] == 0:
                e_start = e[0]
                e_end   = e[1]
                duration = 0
                for j in range(len(activitys[i])):
                    a_start = activitys[i][j][0]
                    a_end   = activitys[i][j][1]   
                    if  a_start <= e_start < a_end <= e_end :
                        duration += a_end - e_start
                    if e_start < a_start < a_end <= e_end  :
                        duration += a_end - a_start               
                    if e_start <= a_start < e_end <= a_end :
                        duration += e_end - a_start         
                    if a_start <= e_start < e_end <= a_end :
                        duration += e_end - e_start
                    
                if 5*f_s>= duration >=20 :  e[2] = 3
                elif 5*f_s < duration :  e[2] = 4
                else : e[2] = 0
        i+=1        
    return event


def data_for_plot(events, start , f_s):
    datas = [[[0,0,0] for _ in range(len(events[i]))] for i in range(len(events))]
    
    for i in range(len(events)):
        for j in range(len(events[i])):
            datas[i][j][0] = start + dt.timedelta(seconds=round(events[i][j][0]/f_s))
            datas[i][j][1] = start + dt.timedelta(seconds=round(events[i][j][1]/f_s))
            datas[i][j][2] = events[i][j][2]
            
    return datas

def data_for_plot_ac(act, start, f_s):
    d_act = [[[0,0,0] for _ in range(len(act[i]))] for i in range(len(act))]
    
    for i in range(len(act)):
        for j in range(len(act[i])):
            d_act[i][j][0] = start + dt.timedelta(microseconds=int(act[i][j][0]/f_s*1000000))
            d_act[i][j][1] = start + dt.timedelta(microseconds=int(act[i][j][1]/f_s*1000000))
            # if d_act[i][j][0] == d_act[i][j][1] : d_act[i][j][1] += dt.timedelta(seconds=1)
            d_act[i][j][2] = round((act[i][j][1] - act[i][j][0])/f_s,2)
            
    return d_act
    

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

def export_plot_data_xml(datas, hypnogram, RWA, channel,path):
    hx = np.array([h[0] for h in hypnogram])
    hy = np.array([h[2] for h in hypnogram])
    
    hy_ticks_labels = ["","MT","N3","N2","N1","R","W","unscored"]

    fig = plt.figure(figsize = (15,10))
    gs = gridspec.GridSpec(nrows= 2, ncols =1, height_ratios=[2,1])
    hyp = plt.subplot(gs[0])
    hyp.step(hx,hy,'royalblue',where= 'post')
    rem = np.ma.masked_where(hy != 5 , hy)
    hyp.step(hx,rem,'r',where= 'post')
    hyp.set_xlim([hx[0],hx[-1]])
    hyp.set_ylim(0,7)
    hyp.set_yticklabels(range(0,7))
    hyp.set_yticklabels(hy_ticks_labels)
    hyp.xaxis.set_major_locator(mdates.HourLocator())
    hyp.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
    hyp.set_ylabel("Stage", fontsize = 10)
    hyp.set_title(channel)
    
    y_ticks_labels = ["","Phasic","Intermediate","Tonic",""]
    ax = range(len(datas))

    x = np.array([d[0] for d in datas])
    y = np.array([d[2] for d in datas])
    t = "Tonic : " + str(round(RWA[0]*100,3)) +" %"
    i = "Intermediate : " + str(round(RWA[1]*100,3)) +" %"
    p = "Phasic : " + str(round(RWA[2]*100,3)) +" %"

    
    ax = plt.subplot(gs[1])
    phasic = np.ma.masked_where(y != 1 , y)
    tonic = np.ma.masked_where(y != 3 , y)
    inter = np.ma.masked_where(y != 2 , y)
    ax.plot(x,tonic, 'bs', markersize = 4.5, label =t )
    ax.plot(x,inter, 'gs', markersize = 4.5, label = i)
    ax.plot(x,phasic, 'ys', markersize = 4.5, label = p)
    

    ax.set_xlim([hx[0],hx[-1]])
    ax.set_yticks(range(0,5), labels = y_ticks_labels)
    colors = ['w','y','g','b','w']
    for ytick, color in zip(ax.get_yticklabels(), colors):
        ytick.set_color(color)
    ax.xaxis.set_major_locator(mdates.HourLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
    ax.grid(axis ='y')
    ax.set_xlabel("Time (h:m:s)", fontsize = 10)
    ax.set_ylabel("Event" , fontsize = 10)
    ax.legend(loc = 'upper left')
    fig.tight_layout()

    for d in datas:
        d[0] = d[0].strftime("%H:%M:%S.%f")
        d[1] = d[1].strftime("%H:%M:%S.%f")
        if   d[2] == 1 : d[2] = 'Phasic'
        elif d[2] == 0 : d[2] = 'Tonic'
        elif d[2] == 2 : d[2] = 'Intermediate'
        elif d[2] == 3 : d[2] = 'Tonic / Phasic'
        elif d[2] == 4 : d[2] = "Tonic / Inter"
        elif d[2] == 11: d[2] = 'Artifact'
        elif d[2] == 10: d[2] = 'No'
        

    df = pd.DataFrame(datas, columns = [ 'Start', 'End', 'Event'])
    with pd.ExcelWriter(path, mode ='a', if_sheet_exists= "overlay", engine = "openpyxl") as writer:
        df.to_excel(writer, sheet_name =channel)
        imgdata = io.BytesIO()
        fig.savefig(imgdata, format ='png')

    wb = xl.load_workbook(path,data_only=True)
    ws = wb[channel]
    im = PILImage.open(imgdata)
    pil_img = im.resize((1200,700))
    pil_img.save("resiged.png")
    xl_img = XLImage("resiged.png")
    ws.add_image(xl_img, "F6")

    wb.save(path)
    wb.close()
    plt.close() 
    gc.collect()
    
def get_ac_RWA(act, epoch,fs):

    ac_RWA = [0 for _ in range(len(epoch))]
    for i in range(len(epoch)):
        tot_rem = 0
        tot_ac  = 0
        for rem in epoch[i] : tot_rem += (rem[1]-rem[0])
        for ac  in act[i] : tot_ac  += ( ac[1]- ac[0])
        ac_RWA[i] = tot_ac/fs/tot_rem
    return ac_RWA


def export_plot_AASM_xml(datas, hypnogram, RWA, channel,path):
    hx = np.array([h[0] for h in hypnogram])
    hy = np.array([h[2] for h in hypnogram])
    
    hy_ticks_labels = ["","MT","N3","N2","N1","R","W","unscored"]

    fig = plt.figure(figsize = (15,10))
    gs = gridspec.GridSpec(nrows= 2, ncols =1, height_ratios=[2,1])
    hyp = plt.subplot(gs[0])
    hyp.step(hx,hy,'royalblue',where= 'post')
    rem = np.ma.masked_where(hy != 5 , hy)
    hyp.step(hx,rem,'r',where= 'post')
    hyp.set_xlim([hx[0],hx[-1]])
    hyp.set_ylim(0,7)
    hyp.set_yticklabels(range(0,7))
    hyp.set_yticklabels(hy_ticks_labels)
    hyp.xaxis.set_major_locator(mdates.HourLocator())
    hyp.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
    hyp.set_ylabel("Stage", fontsize = 10)
    hyp.set_title(channel)
    
    y_ticks_labels = ["","Phasic","Inter","Tonic",""]
    ax = range(len(datas))

    x = np.array([d[0] for d in datas])
    y = np.array([d[2] for d in datas])
    t = "Tonic : " + str(round(RWA[0]*100,3)) +" %"
    p = "Phasic : " + str(round(RWA[1]*100,3)) +" %"
    i = "Any : " + str(round(RWA[2]*100,3)) +" %"

    
    ax = plt.subplot(gs[1])
    phasic = np.ma.masked_where(y != 0 , y)
    tonic = np.ma.masked_where(y != 1 , y)
    any = np.ma.masked_where(y != 2 , y)
    ax.plot(x,tonic+2, 'bs', markersize = 4.5, label =t )
    ax.plot(x,phasic+1, 'ys', markersize = 4.5, label = p)
    ax.plot(x,any, 'gs', markersize = 4.5, label = i)
    

    ax.set_xlim([hx[0],hx[-1]])
    ax.set_yticks(range(0,5), labels = y_ticks_labels)
    colors = ['w','y','g','b','w']
    for ytick, color in zip(ax.get_yticklabels(), colors):
        ytick.set_color(color)
    ax.xaxis.set_major_locator(mdates.HourLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
    ax.grid(axis ='y')
    ax.set_xlabel("Time (h:m:s)", fontsize = 10)
    ax.set_ylabel("Event" , fontsize = 10)
    ax.legend(loc = 'upper left')
    fig.tight_layout()

    for d in datas:
        d[0] = d[0].strftime("%H:%M:%S.%f")
        d[1] = d[1].strftime("%H:%M:%S.%f")
        if   d[2] == 0 : d[2] = 'Tonic'
        elif d[2] == 1 : d[2] = 'Phasic'
        elif d[2] == 2 : d[2] = 'Intermediate'
        elif d[2] == 10 : d[2] = 'No event'
        elif d[2] == 11: d[2] = 'Artifact'

        

    df = pd.DataFrame(datas, columns = [ 'Start', 'End', 'Event'])
    with pd.ExcelWriter(path, mode ='a', if_sheet_exists= "overlay", engine = "openpyxl") as writer:
        df.to_excel(writer, sheet_name =channel)
        imgdata = io.BytesIO()
        fig.savefig(imgdata, format ='png')

    wb = xl.load_workbook(path,data_only=True)
    ws = wb[channel]
    im = PILImage.open(imgdata)
    pil_img = im.resize((1200,700))
    pil_img.save("resiged.png",dpi=(330, 330))
    xl_img = XLImage("resiged.png")
    ws.add_image(xl_img, "F6")

    wb.save(path)
    wb.close()
    plt.close()   
    gc.collect()
    
def get_ac_RWA(act, epoch,fs,art_duration):

    ac_RWA = [0 for _ in range(len(epoch))]
    for i in range(len(epoch)):
        
        total_art = art_duration[i]
        tot_rem = 0
        tot_ac  = 0
        for rem in epoch[i] : 
            tot_rem += (rem[1]-rem[0])
        for ac  in act[i] : tot_ac  += ( ac[1]- ac[0])
        tot_rem -= total_art
        ac_RWA[i] = tot_ac/fs/tot_rem
    return ac_RWA

def make_RWA_metric(activity1,channel,epochs, event, activity2, art_duration):
    RWA_Event = ["Tonic", "Intermediate", "Phasic","Any"]
    ev_index = {
        0 : [0,3,4], # Tonic
        1 : [2,4], # Intermediate
        2 : [1,3], # Phasic
        3 : [0,1,2,3,4] # Any
    }
    percentile = [50, 80, 95]
    RWA_mean  = [[0 for _ in channel] for _ in RWA_Event] 
    RWA_freq  = [[0 for _ in channel] for _ in RWA_Event]
    RWA_score = [[0 for _ in channel] for _ in RWA_Event]
    
    RWA_durations = [[[] for _ in channel] for _ in RWA_Event]
    RWA_percentile = [[[0 for _ in channel] for _ in percentile]for _ in RWA_Event] 

    for j in range(len(activity1)):
        RWA_t = [0 for _ in RWA_Event]
        all_time = 0
        total_art = art_duration[j]
        for ep in epochs[j]:
            all_time += ep[1] - ep[0]
        all_time -= total_art
        all_hour = all_time / 60/60
        for k in range(len(RWA_Event)): # T I P A
            for act in activity1[j]:
                trg = 0
                for ev in event[j]: # 겹치는지
                    s= ev[0]
                    e= ev[1]
                    if ev[2] not in ev_index[k]: continue
                    if act[0]>=s and act[1]<=e:trg =2
                    elif e>=act[0]>=s and act[1]>=e:trg =2
                    elif act[0] <= s and s<=act[1]<=e:trg =2
                    elif act[0] <= s and e<=act[1]:trg =2
                if trg == 2: #겹쳤다
                    if k == 2 and (act[1]- act[0])/200 >5: continue
                    if (k == 0 or k == 1) and (act[1]- act[0])/200 <=5: continue
                    RWA_durations[k][j].append(round((act[1]- act[0])/200,3))
                    RWA_mean[k][j] += (act[1]- act[0])/200 #RWA_sum
                    RWA_t[k] += 1

            for act in activity2[j]:
                trg = 0
                for ev in event[j]: # 겹치는지
                    s= ev[0]
                    e= ev[1]
                    if ev[2] not in ev_index[k]: continue
                    if act[0]>=s and act[1]<=e:trg =2
                    elif e>=act[0]>=s and act[1]>=e:trg =2
                    elif act[0] <= s and s<=act[1]<=e:trg =2
                    elif act[0] <= s and e<=act[1]:trg =2
                if trg == 2: #겹쳤다
                    if k == 2 and (act[1]- act[0])/200 >5: continue
                    if (k == 0 or k == 1) and (act[1]- act[0])/200 <=5: continue
                    RWA_durations[k][j].append(round((act[1]- act[0])/200,3))
                    RWA_mean[k][j] += (act[1]- act[0])/200 #RWA_sum
                    RWA_t[k] += 1
                           
            if RWA_t[k] != 0:
                
                RWA_score[k][j] = RWA_mean[k][j] / all_time * 100
                RWA_score[k][j] = round(RWA_score[k][j],3) 
                
                    
                RWA_mean[k][j] /= RWA_t[k]
                RWA_mean[k][j] = round(RWA_mean[k][j],3)
                
                RWA_freq[k][j] = RWA_t[k]/ all_hour
                RWA_freq[k][j] = round(RWA_freq[k][j], 3)    
                
                if all_time < 600:
                    RWA_score[k][j]  = 0     
                    RWA_mean[k][j]  = 0
                    RWA_freq[k][j]  = 0          
    
    for k in range(len(RWA_Event)):
        for p in range(len(percentile)):
            for j in range(len(activity1)):
                std_act = sorted(RWA_durations[k][j])
                n_percent = math.floor(len(std_act)*percentile[p]/100)
                if len(std_act) == 0:
                    RWA_percentile[k][p][j] = 0
                else:
                    RWA_percentile[k][p][j] = std_act[n_percent]
        
                if all_time < 600:
                    RWA_score[k][p][j]  = 0     
                    RWA_mean[k][p][j]  = 0
                    RWA_freq[k][p][j]  = 0 
    return RWA_mean , RWA_freq, RWA_score, RWA_percentile

def write_CRWA(path, hypnogram, RWA_score,RWA_mean,RWA_percentile,RWA_freq,d_act, AHI, channel):
    result_path = path+'/comb_CRWA.xlsx'
    wb = xl.Workbook()
    wb.save(result_path)
    wb.close()
    for i in range(len(d_act)):
        export_plot_act_data_xml(d_act[i], hypnogram, RWA_score[-1][i], channel[i],result_path)

            
    gc.collect()
    wb = xl.load_workbook(result_path,data_only=True)
    wb.remove(wb['Sheet'])
    ws = wb.create_sheet("RWA summary")
    ws["A1"] = "( % )"
    ws["H1"] = "REM AHI"
    ws["H2"] = AHI
    
    for p in range(len(para)):
        ind = "A"+str(p+2)
        ws[ind] = para[p]
    for j in range(len(RWA_score[0])):
        r = 2
        ws.cell(1, j+2).value = channel[j]
        for e in range(len(RWA_Event)):
            ws.cell(r,j+2).value = RWA_score[e][j]; r+=1
            ws.cell(r,j+2).value = RWA_mean[e][j] ; r+=1
            for p in range(len(percentile)):
                ws.cell(r,j+2).value = RWA_percentile[e][p][j] 
                r+=1
            ws.cell(r,j+2).value = RWA_freq[e][j]
            r+=1
    wb.save(result_path)
    wb.close()   
    gc.collect()
    
    
def export_plot_act_data_xml(act, hypnogram, RWA_ac, ch,path):
    hx = np.array([h[0] for h in hypnogram])
    hx_e = np.array([h[1] for h in hypnogram])
    hy = np.array([h[2] for h in hypnogram])
    hy_ticks_labels = ["","MT","N3","N2","N1","R","W","unscored"]

    fig = plt.figure(figsize = (15,10))
    gs = gridspec.GridSpec(nrows= 2, ncols =1, height_ratios=[2,1])
    hyp = plt.subplot(gs[0])
    hyp.step(hx,hy,'royalblue',where= 'post')
    rem = np.ma.masked_where(hy != 5 , hy)
    hyp.step(hx,rem,'r',where= 'post')
    hyp.set_xlim([hx[0],hx[-1]])
    hyp.set_ylim(0,7)
    hyp.set_yticklabels(range(0,7))
    hyp.set_yticklabels(hy_ticks_labels)
    hyp.xaxis.set_major_locator(mdates.HourLocator())
    hyp.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
    hyp.set_ylabel("Stage", fontsize = 10)
    hyp.set_title(ch)
    
    y_ticks_labels = ["","activity",""]
    ax = range(len(act))

    x = np.array([d[0] for d in act])
    t = "activity : " + str(round(RWA_ac,3)) +" %"


    ax = plt.subplot(gs[1])
    y = np.ones(len(x))
    ax.plot(x,y, 'gs', markersize = 1.5, label = t)
    

    ax.set_xlim([hx[0],hx[-1]])
    ax.set_yticks(range(0,3), labels = y_ticks_labels)
    colors = ['w','g','w']
    for ytick, color in zip(ax.get_yticklabels(), colors):
        ytick.set_color(color)
    ax.xaxis.set_major_locator(mdates.HourLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
    ax.grid(axis ='y')
    ax.set_xlabel("Time (h:m:s)", fontsize = 10)
    ax.set_ylabel("Event" , fontsize = 10)
    ax.legend(loc = 'upper left')
    fig.tight_layout()

    for d in act:
        d[0] = d[0].strftime("%H:%M:%S.%f")
        d[1] = d[1].strftime("%H:%M:%S.%f")
        

    df = pd.DataFrame(act, columns = ['Start', 'End', 'Duration'])
    with pd.ExcelWriter(path, mode ='a', if_sheet_exists= "overlay", engine = "openpyxl") as writer:
        df.to_excel(writer, sheet_name =ch)
        imgdata = io.BytesIO()
        fig.savefig(imgdata, format ='png')

    wb = xl.load_workbook(path,data_only=True)
    ws = wb[ch]
    im = PILImage.open(imgdata)
    pil_img = im.resize((1200,700))
    pil_img.save("resiged.png",dpi=(330, 330))
    xl_img = XLImage("resiged.png")
    ws.add_image(xl_img, "F6")

    wb.save(path)
    wb.close()  
    plt.close()
    gc.collect()
    
def export_plot_RAI_xml(act, hypnogram, RAI, ch,path):
    hx = np.array([h[0] for h in hypnogram])
    hy = np.array([h[2] for h in hypnogram])
    hy_ticks_labels = ["","MT","N3","N2","N1","R","W","unscored"]

    fig = plt.figure(figsize = (30,50))
    gs = gridspec.GridSpec(nrows= 2, ncols =1, height_ratios=[2,6])
    hyp = plt.subplot(gs[0])
    hyp.step(hx,hy,'royalblue',where= 'post')
    rem = np.ma.masked_where(hy != 5 , hy)
    hyp.step(hx,rem,'r',where= 'post')
    hyp.set_xlim([hx[0],hx[-1]])
    hyp.set_ylim(0,7)
    hyp.set_yticklabels(range(0,7))
    hyp.set_yticklabels(hy_ticks_labels)
    hyp.xaxis.set_major_locator(mdates.HourLocator())
    hyp.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
    hyp.set_ylabel("Stage", fontsize = 30)
    hyp.set_title(ch)
    
    y_ticks_labels = ["","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20",""]
    ax = range(len(act))

    for i in range(1,21):
        x = np.array([d[0] for d in act if d[2] == i])

        ax = plt.subplot(gs[1])
        y = [ i for _ in x]
        ax.plot(x,y, 'gs', markersize = 3)
    
    ax.set_xlim([hx[0],hx[-1]])
    ax.set_yticks(range(0,22), labels = y_ticks_labels)
    colors = ['w'] +['g'] *20 +['w']
    for ytick, color in zip(ax.get_yticklabels(), colors):
        ytick.set_color(color)
    ax.xaxis.set_major_locator(mdates.HourLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
    ax.grid(axis ='y')
    ax.set_xlabel("Time (h:m:s)", fontsize = 30)
    ax.set_ylabel("Amplitude (uV)" , fontsize = 30)
    ax.legend(loc = 'upper left')
    fig.tight_layout()

    for d in act:
        d[0] = d[0].strftime("%H:%M:%S.%f")
        d[1] = d[1].strftime("%H:%M:%S.%f")
        

    df = pd.DataFrame(act, columns = ['Start', 'End', 'uV'])
    with pd.ExcelWriter(path, mode ='a', if_sheet_exists= "overlay", engine = "openpyxl") as writer:
        df.to_excel(writer, sheet_name =ch)
        imgdata = io.BytesIO()
        fig.savefig(imgdata, format ='png')

    wb = xl.load_workbook(path,data_only=True)
    ws = wb[ch]
    im = PILImage.open(imgdata)
    pil_img = im.resize((1800,1800))
    pil_img.save("resiged.png",dpi=(330, 330))
    xl_img = XLImage("resiged.png")
    ws.add_image(xl_img, "F6")

    wb.save(path)
    wb.close()  
    plt.close()
    gc.collect()
     
def write_SINBAR(path, hypnogram, d_event, channel, RWA):
    result_path = path+'/comb_SINBAR.xlsx'
    wb = xl.Workbook()
    wb.save(result_path)
    wb.close()
    for i in range(len(d_event)):
        rwa = [0,0,0,0]
        for j in range(4):
            rwa[j] = RWA[i][j]
        
        export_plot_data_xml(d_event[i], hypnogram, rwa, channel[i],result_path)
    gc.collect()
    wb = xl.load_workbook(result_path,data_only=True)
    wb.remove(wb['Sheet'])
    ws = wb.create_sheet("RWA summary")
    ws["A1"] = "( % )"
    ws["A2"] ="Tonic"
    ws["A3"] ="Intermediate"
    ws["A4"] ="Phasic"
    ws["A5"] ="Any"
    
    for i in range(4):
        for j in range(len(RWA)):
            ws.cell(1   , j+2).value = channel[j]
            ws.cell(i+2 , j+2).value = round(RWA[j][i]*100,3)
    
    wb.save(result_path)
    wb.close()
    gc.collect()
    
def write_AASM(path, hypnogram, d_event, channel, RWA):
    result_path = path+'/comb_AASM.xlsx'
    wb = xl.Workbook()
    wb.save(result_path)
    wb.close()
    for i in range(len(d_event)):
        rwa = [0,0,0]
        for j in range(3):
            rwa[j] = RWA[i][j]
        
        export_plot_AASM_xml(d_event[i], hypnogram, rwa, channel[i],result_path)
    gc.collect()
    wb = xl.load_workbook(result_path,data_only=True)
    wb.remove(wb['Sheet'])
    ws = wb.create_sheet("RWA summary")
    ws["A1"] = "( % )"
    ws["A2"] ="Tonic"
    ws["A3"] ="Phasic"
    ws["A4"] ="Any"
    
    for i in range(3):
        for j in range(len(RWA)):
            ws.cell(1   , j+2).value = channel[j]
            ws.cell(i+2 , j+2).value = round(RWA[j][i]*100,3)
    
    wb.save(result_path)
    wb.close()
    gc.collect()

def write_RAI(path, hypnogram, d_event, channel, RAI):
    result_path = path+'/comb_RAI.xlsx'
    wb = xl.Workbook()
    wb.save(result_path)
    wb.close()
    for i in range(len(d_event)):
        
        export_plot_RAI_xml(d_event[i], hypnogram, RAI[i], channel[i],result_path)
    gc.collect()
    wb = xl.load_workbook(result_path,data_only=True)
    wb.remove(wb['Sheet'])
    ws = wb.create_sheet("RWA summary")
    ws["A1"] = "( % )"
    ws["A2"] ="RAI"

    for j in range(len(RAI)):
        ws.cell(1   , j+2).value = channel[j]
        ws.cell(2 , j+2).value = round(RAI[j],3)
    
    wb.save(result_path)
    wb.close()
    gc.collect()
    



def make_RWA(event, channel, RWA_AASM): #ch epoch event
    RWA = [[0,0,0,0] for _ in channel]
    
    for i in range(len(channel)):
        art_count = np.sum(np.array(event[i]) == 11)
        for j in range(4):
            ev_ind = ev_index[j]
            for ev in event[i]:
                if ev[2] in ev_ind:
                    RWA[i][j] +=1
            RWA[i][j] /=(len(event[i]) - art_count)
            if j == 0:
                RWA[i][j] = RWA_AASM[i][j]            
    return RWA

def make_RWA_AASM(event, channel): #ch epoch event
    RWA = [[0,0,0] for _ in channel]
    
    for i in range(len(channel)):
        art_count = np.sum(np.array(event[i]) == 11)
        for j in range(3):
            ev_ind = ev_index_AASM[j]
            for ev in event[i]:
                if ev[2] in ev_ind:
                    RWA[i][j] +=1
            RWA[i][j] /=(len(event[i]) - art_count)
            
    return RWA

def make_event3(event):
    events3  = [np.empty((0,3),int) for _ in range(len(event))]
    
    for i in range(len(event)):
        for j in range(len(event[i])):
            if event[i][j][2] != 10:
                events3[i] = np.append(events3[i],np.array([[event[i][j][0],event[i][j][1],event[i][j][2]]]),axis =0 )
                
    return events3
    

def make_RAI(EMGs, f_s, epochs, artifacts):
    bouts = int(1*f_s)
    rais = [1000 for _ in epochs]
    activity_RA = []
    
    for i in range(len(epochs)):
        es = []
        RA = []
        for j in range(len(epochs[i])):
            start = epochs[i][j][0]*f_s
            end = epochs[i][j][1]*f_s
            t_end = start + bouts
            
            while t_end <= end: 
                ##############################
                ## artifact 제외
                art_t = 0
                for art in artifacts[i]:
                    if start <= art[0]*f_s <= t_end:
                        art_t= 1; break
                    elif art[0]*f_s <= start <= art[1]*f_s:
                        art_t= 1; break
                
                if art_t == 1:
                    RA.append([start, t_end, -1])       
                    t_end += bouts
                    start += bouts     
                    continue
                ###############################
                ## 1s mini-epoch 평균 값에서 baseline 값 뺌
                e = np.mean(np.abs((EMGs[i][start: t_end])))
                st = start - 30*bouts
                ed = st + bouts
                bases = []
                while ed <= t_end+30*bouts:
                    bases.append(np.mean(np.abs((EMGs[i][st: ed]))))
                    st += bouts
                    ed += bouts
                e -= np.min(np.array(bases)) # baseline
                ###############################
                ## uV에 따라 분리
                if e < 20e-06: 
                    es.append(int(e // 1e-06)+1)
                    RA.append([start, t_end, int(e//1e-06)+1])
                else :
                    es.append(20)
                    RA.append([start, t_end, 20])
                ##############################
                t_end += bouts
                start += bouts
        activity_RA.append(np.array(RA))
        es = np.array(es)
        ## 1uV 이하 값의 비율 == RAI
        art_num = np.sum(es == 2) + np.sum(es == -1)
        rais[i] = np.sum(es == 1) / (len(es) - art_num)
    
    a= 1
    return np.array(activity_RA), rais

def get_art_duration(artifacts,epochs):
    dur = [0,0,0,0,0]
    for i in range(3):
        for epoch in epochs[i]:
            start = epoch[0]
            t_end = start + 1
            R = [start, t_end]
            end = epoch[1]
            while t_end <= end:
                trg = 0
                for art in artifacts[i]:
                    if art[0] < R[0] < R[1] < art[1]: trg = 1
                    elif art[0] < R[0] < art[1] < R[1] :  trg = 1
                    elif R[0] < art[0] < R[1] < art[1] :  trg = 1
                    elif R[0] < art[0] < art[1] < R[1] :  trg = 1
                if trg == 1:
                    dur[i] += 1
                start += 1
                t_end += 1
                R = [start, t_end]
        a = 1
    return dur
        

RWA_Event = ["Any","Phasic", "Intermediate", "Tonic"]
percentile = [50, 80, 95]
ev_index = {
    0 : [0,3,4], # Tonic
    1 : [2,4], # Intermediate
    2 : [1,3], # Phasic
    3 : [0,1,2,3,4] # Any
}

ev_index_AASM = {
    0 : [0], # Tonic
    1 : [1], # Phasic
    2 : [0,1,2], # Any
}

para = [
    "RWAC_T","DURmean_T","p50_T","p80_T","p95_T","RWAI_T",
    "RWAC_I","DURmean_I","p50_I","p80_I","p95_I","RWAI_I",
    "RWAC_P","DURmean_P","p50_P","p80_P","p95_P","RWAI_P",
    "RWAC_A","DURmean_A","p50_A","p80_A","p95_A","RWAI_A",
]
def merge_arts(A,B):
    # 모든 이벤트를 하나의 리스트로 합침
    A_start = []; A_end = []
    for a in A:
        A_start.append(a[0])
        A_end.append(a[1])

    B_start = []; B_end = []
    for a in B:
        B_start.append(a[0])
        B_end.append(a[1])
        
    events = []
    for start, end in zip(A_start, A_end):
        events.append((start, 'start'))
        events.append((end, 'end'))
    for start, end in zip(B_start, B_end):
        events.append((start, 'start'))
        events.append((end, 'end'))
    
    # 이벤트를 시간순으로 정렬
    events.sort()
    
    merged = []
    
    active_events = 0
    current_start = None
    
    for time, type in events:
        if type == 'start':
            if active_events == 0:
                current_start = time
            active_events += 1
        elif type == 'end':
            active_events -= 1
            if active_events == 0:
                merged.append([current_start, time])
    
    return np.array(merged)
def merge_events(A,B):
    # 모든 이벤트를 하나의 리스트로 합침
    A_start = []; A_end = []
    for a in A:
        A_start.append(a[0])
        A_end.append(a[1])

    B_start = []; B_end = []
    for a in B:
        B_start.append(a[0])
        B_end.append(a[1])
        
    events = []
    for start, end in zip(A_start, A_end):
        events.append((start, 'start'))
        events.append((end, 'end'))
    for start, end in zip(B_start, B_end):
        events.append((start, 'start'))
        events.append((end, 'end'))
    
    # 이벤트를 시간순으로 정렬
    events.sort()
    
    merged = []
    
    active_events = 0
    current_start = None
    
    for time, type in events:
        if type == 'start':
            if active_events == 0:
                current_start = time
            active_events += 1
        elif type == 'end':
            active_events -= 1
            if active_events == 0:
                merged.append([current_start, time,1])
    
    return np.array(merged)

def combine_act(acts,triggers):
    res = [[] for _ in [0,1,2,3,4]]
    for i in [0,1,2]:
        if i == 0: res[i] = np.array(acts[0]); continue
        elif i ==1: 
            a = acts[1]; b = acts[2]
        elif i ==2:
            a = acts[3]; b = acts[4]
        
        res[i] = merge_events(a,b)
        
    for i in [3,4]:
        if i == 3: a = res[0]; b = res[2]
        elif i ==4: 
            a = res[3]; b = res[1]
        
        res[i] = merge_events(a,b)     
    
    for i in range(5):
        if triggers[i] == False:
            res[i] = np.array([])   
    return res

def combine_art(acts,triggers):
    res = [[] for _ in [0,1,2,3,4]]
    for i in [0,1,2]:
        if i == 0: res[i] = np.array(acts[0]); continue
        elif i ==1: 
            a = acts[1]; b = acts[2]
        elif i ==2:
            a = acts[3]; b = acts[4]
        
        res[i] = merge_arts(a,b)
        
    for i in [3,4]:
        if i == 3: a = res[0]; b = res[2]
        elif i ==4: 
            a = res[3]; b = res[1]
        
        res[i] = merge_arts(a,b)     
    
    for i in range(5):
        if triggers[i] == False:
            res[i] = np.array([])   
    return res

def combine_rai(acts, rais):
    res = [[] for _ in [0,1,2,3,4]]
    rai = [0 for _ in [0,1,2,3,4]]
    for i in [0,1,2,3,4]:
        if i == 0: res[i] = acts[0]; rai[i] = rais[i]; continue
        elif i ==1: 
            a = acts[1]; b = acts[2]
        elif i ==2:
            a = acts[3]; b = acts[4]
        
        res[i], rai[i] = merge_rai(a,b)

    for i in [3,4]:
        if i == 3: a = res[0]; b = res[2]
        elif i ==4: 
            a = res[3]; b = res[1]
        
        res[i], rai[i] = merge_rai(a,b)      
    return res, np.array(rai)

def merge_rai(a,b):
    res = []
    for i in range(len(a)):
        if -1 in [a[i][2],b[i][2]]: res.append([a[i][0],a[i][1],-1])
        elif a[i][2] >= b[i][2]: res.append(a[i])
        else: res.append(b[i])
    res= np.array(res)
    art_num = np.sum(res == 2) + np.sum(res == -1)
    rai = np.sum(res == 1) / (len(res) - art_num)
    return res, rai
act_duration = [[],[],[],[],[],[]]
mean_duration = [[],[],[],[],[],[]]
num_subjects = [ 0, 0, 0 ,0 ,0]


def comb_event(events,triggers): #ch ep [s,e,t]
    res = [[] for _ in [0,1,2,3,4]]
    for i in [0,1,2]:
        if triggers[i] == False:
            res[i] = np.array([]); continue
        if i == 0: res[i] = np.array(events[0]); continue
        elif i ==1: 
            a = events[1][:]; b = events[2][:]
        elif i ==2:
            a =events[3][:]; b = events[4][:]
        
        res[i] = merge_ev(a,b)
        
    for i in [3,4]:
        if triggers[i] == False:
            res[i] = np.array([]); continue
        if i == 3: a = res[0][:]; b = res[2][:]
        elif i ==4: 
            a = res[3][:]; b = res[1][:]
        
        res[i] = merge_ev(a,b)     
  
    return res    

def merge_ev(A, B):
    res = []
    for i in range(len(A)):
        t = -1
        for j in range(len(B)):
            if A[i][0] == B[j][0]:
                t = j
        if t == -1:
            continue

        # Create a shallow copy to ensure the original A[i] is not modified
        tt = A[i][:]  # This is a shallow copy

        if (1 in [A[i][2], B[t][2]] or 2 in [A[i][2], B[t][2]] or 10 in [A[i][2], B[t][2]]) and 0 in [A[i][2], B[t][2]]:
            tt[2] = 0
            res.append(tt)
        elif 11 in [A[i][2], B[t][2]]:
            tt[2] = 11
            res.append(tt)
        elif 10 in [A[i][2], B[t][2]]:
            if A[i][2] < 5:
                tt[2] = A[i][2]
            if B[t][2] < 5:
                tt[2] = B[t][2]
            res.append(tt)
        else:
            tt[2] = max(A[i][2], B[t][2])
            res.append(tt)
    
    return np.array(res)
                
