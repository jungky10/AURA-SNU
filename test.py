import mne


path = 'E:/PSG_data/00_PSG/R001/Traces.edf'

data = mne.io.read_raw_edf(path, preload= True)

# data.plot()

print(data.info)
input()
