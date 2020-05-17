import pandas as pd
import openpyxl as op
import matplotlib.pyplot as plt
def userdetails():
#df = pd.read_excel(r"File_Location")
    df = pd.read_excel(r"D:\crp\CropsDataFile (1).xlsx") #Enter thefile location in your PC
    print(df)
    """for i in df.index:
        x=df['Crop_Year'][i]
        y=df['Production'][i]
        plt.barh(x,y)
    plt.show()"""
def managerchanges():

    wb = op.load_workbook("D:\crp\CropsDataFile (1).xlsx") #Enterthe file location in your PC
    ws = wb['Sheet1']
    n = int(input("Enter The Number Of Crops You Want To Add: "))
    for i in range(0,n):
        Col1 = input("\nEnter The State Name: ")
        Col2 = input("\nEnter The District Name: ")
        Col3 = int(input("\nEnter The Crop Year: "))
        Col4 = input("\nEnter The Season: ")
        Col5 = input("\nEnter The Crop Name: ")
        Col6 = int(input("\nEnter The Area: "))
        Col7 = int(input("\nEnter The Amount Of Production: "))
        ws.append([Col1, Col2, Col3, Col4, Col5, Col6, Col7])
    wb.save("D:\crp\CropsDataFile (1).xlsx") #Enter the filelocation in your PC
    wb.close()
    df = pd.read_excel(r"D:\crp\CropsDataFile (1).xlsx") #Enter thefile location in your PC
    print(df)
print("Enter either user or manager: ")
name = input()
if(name == "user" or name == "USER" or name == "User"):
    userdetails()
elif(name == "manager" or name == "MANAGER" or name == "Manager"):
    print("Enter The Manager's Password: ")
    password = int(input())
    if(password == 9876): #password is 9876
        managerchanges()
    else:
        print("Wrong Password!!!")
else:
    pass