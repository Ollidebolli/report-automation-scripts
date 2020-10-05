import pandas as pd
import glob

frames = []
for IR in glob.glob('*IR' + '*csv'):
    frames.append(pd.read_csv(IR))

pd.concat(frames,ignore_index=False).reset_index(drop=True).to_csv('Presales Investment - Employee Details.csv')