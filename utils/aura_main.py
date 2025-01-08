from .helper_functions import *

def AURA_main(data,path,f_s,data_start,channels,data_length,event_path) :
    try:
        [REM, start_index, event_start, hypnogram,format] = make_stage_event(event_path, data_start)
        [artifacts, AHI] = make_artifact(event_path,event_start,format, REM)
    except:
        try:
            [REM, start_index, event_start, hypnogram,format] = make_stage_event2(event_path, data_start)
            [artifacts, AHI] = make_artifact2(event_path,event_start,format, REM)  # channel x [start, end]  
        except:
            [REM, start_index, event_start, hypnogram,format] = make_stage_event3(event_path, data_start)
            [artifacts, AHI] = make_artifact3(event_path,event_start,format, REM)  # channel x [start, end]        
  
    print("REM AHI:",AHI)
    print("Filtering...")
    f_s = 200
    [EMG, channel,triggers] = filter_Seperate(data, start_index, channels)
    print("EMG channel:",channel)
    [epochs, art_epochs] = make_REM_epochs(REM, artifacts)
    print("Making RMS...")
    rms = make_rms(EMG,epochs,f_s)
    print("Making REM Baseline...")
    baseline = make_baseline(rms, epochs, art_epochs)
    print("getting RWA...")
    activity1,_ = make_activity(baseline, artifacts, EMG, f_s, epochs) ## [Activity, % of line noise artifacts]
    #################################################### 
    channel = ['MT','AT','FDS','MT+FDS',"MT+AT+FDS"]
    activity_RA, RAIs = make_RAI(EMG, f_s, epochs, artifacts)
    activity_RA, RAIs = combine_rai(activity_RA, RAIs)
    print("Making RWA Event...")
    event, event30  = make_event(activity1, epochs, f_s, artifacts) # 30s epoch 이벤트, 3s 이벤트 분리
    print("getting Tonic RWA...")
    activity2 = make_activity2(EMG, f_s,event)
    print("Making RWA Tonic Event...")
    event = make_event2(activity2, f_s, event) # 0: tonic, 1: phasic, 2: inter, 3: t/p, 4: t/i
    event = comb_event(event, triggers)
    event30 = comb_event(event30, triggers)
    
    RWA_AASM = make_RWA_AASM(event30, channel) #
    RWA_SINBAR = make_RWA(event, channel, RWA_AASM) #
    event = make_event3(event)
    print("Making RWA CRWA file...") 
    activity1 = combine_act(activity1,triggers)
    activity2 = combine_act(activity2,triggers)
    artifacts = combine_art(artifacts, triggers)
    art_duration = get_art_duration(artifacts, epochs)

    d_act = data_for_plot_ac(activity1, event_start, f_s)
    RWA_mean , RWA_freq, RWA_score, RWA_percentile = make_RWA_metric(activity1, channel,epochs, event, activity2, art_duration) #
    write_CRWA(path, hypnogram, RWA_score,RWA_mean,RWA_percentile,RWA_freq, d_act, AHI, channel)

    print("Making RWA SINBAR event file...")
    d_event = data_for_plot(event, event_start, f_s)
    write_SINBAR(path, hypnogram, d_event, channel, RWA_SINBAR) 

    print("Making RWA AASM event file...")
    d_event = data_for_plot(event30, event_start, f_s)
    write_AASM(path, hypnogram, d_event, channel, RWA_AASM)
    
    print("Making RWA RAI event file...")
    d_event = data_for_plot(activity_RA, event_start, f_s)
    write_RAI(path, hypnogram, d_event, channel, RAIs)