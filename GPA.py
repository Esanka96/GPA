import pandas as pd
df = pd.read_excel("Modeule_file.xlsx", sheet_name='Sheet1')
list1 = list(df['Number of Semester'])
N=int(list1[0])
list3=[]
list4=[]
X=1
table=[]
Sem_count=[]

for i in range(N):
    Number=str(i+1)
    Sem1='Sem'+Number
    Weight1='Weight'+Number
    list1 = list(df[Sem1])
    list2=list(df[Weight1])
    list3.append(list1)
    list4.append(list2)
    Sem_count.append(Sem1)

List1_of_Grade=[]
List1_of_Weigth=[]
for i in range(len(list3)):
    List2_of_Grade=[]
    List2_of_Weigth=[]
    for j in range(len(list3[i])):
        if str(list3[i][j])=='nan':
            break
        List2_of_Grade.append(str(list3[i][j]))
        List2_of_Weigth.append(str(list4[i][j]))
    List1_of_Grade.append(List2_of_Grade)
    List1_of_Weigth.append(List2_of_Weigth)

List1_of_Credit=[]
for i in range(len(List1_of_Grade)):
    List2_of_Credit=[]
    for j in range(len(List1_of_Grade[i])):
        listResult=["A+","A","A-","B+","B","B-","C+","C","C-","D","I-WE"]
        listCredit=[4.2,4,3.7,3.3,3,2.7,2.3,2,1.5,1,0]
        for k in range(len(listResult)):
            if listResult[k]==List1_of_Grade[i][j]:
                Credit=listCredit[k]
        List2_of_Credit.append(Credit)
    List1_of_Credit.append(List2_of_Credit)

if X==1:
    List1_GPA=[]
    List1_Values=[]
    for i in range(len(List1_of_Weigth)):
        Total_Weight=0
        Total_Amount=0
        for j in range(len(List1_of_Weigth[i])):
            Total_Amount=Total_Amount+(float(List1_of_Weigth[i][j])*float(List1_of_Credit[i][j]))
            Total_Weight=Total_Weight+float(List1_of_Weigth[i][j])
        GPA=round(Total_Amount/Total_Weight,2)
        List1_GPA.append(GPA)
        List1_Values.append(Total_Weight)

    Total_Value=0
    Total_Weight=0
    List_of_final_GPA=[]
    for i in range(len(List1_Values)):
        Total_Value=Total_Value+(List1_Values[i]*List1_GPA[i])
        Total_Weight=Total_Weight+List1_Values[i]
        Final_GPA=round(Total_Value/Total_Weight,2)
        List_of_final_GPA.append(Final_GPA)
    
table.append(Sem_count)
table.append(List1_GPA)
table.append(List_of_final_GPA)
from tabulate import tabulate
print(tabulate(table))

import pandas as pd
df = pd.DataFrame({'Sems': Sem_count,'Sem GPA': List1_GPA,'Final GPA' :List_of_final_GPA })
writer = pd.ExcelWriter('Final.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.close()
