#Data entry Application
import pandas as pd

#Get format inputs from Excel file
filename = input("Enter filename: ")
df = pd.read_excel(filename)
print(df)

#Write the data into Word file
f = open('newfile.doc','w')
line = input("Enter the header: ")
f.write(line)
f.close()
nump = int(input("Enter no. of visits: "))
i=0
valN = input('Enter N: ')
c=1
for c in range (1, nump+1):
    f = open('newfile.doc', 'a')
    col = df[c]
    valM = col[0]
    valSD = col[1]
    valMed = col[2]
    valMin = col[3]
    valMax = col[4]
    i+=1
    f.write("\nVisit"+str(i)+"\nN "+valN+"\nMean "+str(valM)+" ("+str(valSD)+")"+
              "\nMedian "+str(valMed)+"\nMin,Max "+str(valMin)+","+str(valMax))
    f.close()
    c+=1
print("newfile.doc created. Kindly check!")
