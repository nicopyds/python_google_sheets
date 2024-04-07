import pandas as pd

def create_dummy_xslx_file():
    
    data1 = [1, 2, 3]
    data2 = [3, 4, 5]
    data3 = [6, 7, 8]
    
    df1 = pd.DataFrame(data1).rename(columns={0:"folder_1"})
    df2 = pd.DataFrame(data2).rename(columns={0:"folder_2"})
    df3 = pd.DataFrame(data3).rename(columns={0:"folder_3"})

    with pd.ExcelWriter("agg.xlsx") as writer:
        df1.to_excel(writer, sheet_name="folder_1", index=False)
        df2.to_excel(writer, sheet_name="folder_2", index=False)
        df3.to_excel(writer, sheet_name="folder_3", index=False)