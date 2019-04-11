import os
 
def file_name(file_dir): 
    L=[] 
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if os.path.splitext(file)[1] == '.pdf':
                L.append(os.path.join(root, file))
    return L

'''
import os
 
def file_name(file_dir): 
    for root, dirs, files in os.walk(file_dir):
        return files
'''
a = file_name('/Users/Kang/Downloads/Yuanta_ETF')

import pandas as pd
df = pd.DataFrame(a)
df.to_csv('/Users/Kang/Desktop/file_path.csv', index = False, encoding = 'big5hkscs')
