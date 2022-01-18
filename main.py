import pandas as pd
import openpyxl


pd.set_option('display.max_columns', None)
# laDrops_df = pd.DataFrame(columns=headers)
file = 'Liberal Arts.csv'
laEnrollment_df = pd.read_csv(file)
laLecture = laEnrollment_df['Instruction Mode Description'] != 'Laboratory'
lectureEnrollment_df = laEnrollment_df[laLecture]


# print(lectureEnrollment_df)
Engl100Drops = lectureEnrollment_df['Course'] == 'English 100'
Engl100Drops_df = lectureEnrollment_df[Engl100Drops]
Engl100Drops_df = Engl100Drops_df.reset_index(drop=True)
Engl100Drops_df.sort_values(by=['Employee ID', 'Enrollment Add Date'], inplace=True)
Engl100Drops_df = Engl100Drops_df.reset_index(drop=True)
print(Engl100Drops_df)
Engl100Drops_df['Enrollment Drop Date'] = Engl100Drops_df['Enrollment Drop Date'].fillna(0)


multiplesList = []
dropIndexList = []
for i in range(len(Engl100Drops_df)-1):
    if Engl100Drops_df.loc[i, 'Employee ID'] == Engl100Drops_df.loc[i+1, 'Employee ID']:
        Engl100Drops_df.drop(index=i)
        multiplesList.append(Engl100Drops_df.loc[i, 'Employee ID'])


for i in range((len(Engl100Drops_df)-1),-1,-1):
    for number in multiplesList:
        if number == Engl100Drops_df.loc[i, 'Employee ID']:
            if Engl100Drops_df.loc[i, 'Enrollment Drop Date'] != 0:
                # print(i, number, Engl100Drops_df.loc[i, 'Employee ID'])
                if i not in dropIndexList:
                    dropIndexList.append(i)
    # print(dropIndexList)
    update_df = Engl100Drops_df.drop(dropIndexList)

update_df = update_df.loc[update_df['Enrollment Drop Date'] != 0]
update_df = update_df.reset_index()
# update_df.index.name = 'y'
update_df.to_excel('data.xlsx')

EOPSstudents = pd.read_csv('EOPS.csv')

for j in range(len(EOPSstudents)):
    for i in range(len(update_df) - 1):
        if EOPSstudents.loc[j, 'ID'] == update_df.loc[i, 'Employee ID']:
            update_df.loc[i, 'EOPS'] = 'EOPS'




# laEnrollment_df.dropna(subset=['Enrollment Drop Date'], inplace=True)

update_df.to_excel('data.xlsx')
Engl100Drops_df.to_excel('DropsEng100.xlsx')
