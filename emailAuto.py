import os 
import pandas as pd
import smtplib
from email.message import EmailMessage
from getpass import getpass
import win32com.client as win32
import imghdr


# Reads csv file of 
df=pd.read_csv("email.csv")
receivers_email=df["Email"].values #Email column from CSV file will be read
sub=("Test Mail ") #Enter subject of your mail

name=df["Fname"].values #First name column from CSV file will be read




zipped=zip(receivers_email,name)

for(a,c) in zipped: #a is recievers_email c is name
    
    # Open the Outlook
    outlook = win32.Dispatch('outlook.application')
    
    # Create the email
    msg = outlook.CreateItem(0)
    files=[(r"pythonpdf.pdf")]
    
    for file in files:
        
        with open(file,'rb') as f:
            
            file_data=f.read()
            file_name=f.name
            
       
        msg.To=a
        msg.Subject=sub
        
        #Body of the mail
        msg.HTMLBody=f"""
        Hi {c},
         <p>Lorem ipsum dolor sit amet, consectetur 
         adipisicing elit. Doloremque ipsam eaque repellendus quam sequi 
         neque quis numquam tempora voluptates reprehenderit. Magnam, harum aliquam 
         delectus culpa quasi labore? A fugiat nulla ea dolor, reiciendis voluptatem voluptas 
         delectus voluptatum unde exercitationem rem quibusdam veritatis ipsa, 
         aperiam iste sit? Laudantium, voluptas! Amet unde nostrum quibusdam similique veniam, 
         voluptates inventore animi culpa voluptatem optio nihil molestiae possimus error repellat 
         ducimus pariatur cupiditate voluptatum fugiat! Sit illum facilis quod voluptate molestiae, 
         consectetur, praesentium iure cumque unde veniam doloribus exercitationem neque dolor dolorum 
         accusantium sed magni quibusdam sapiente non hic, odit quasi corporis. Quia rem dignissimos, 
         cum nulla asperiores non deleniti laboriosam aperiam eaque corrupti aliquid repellendus unde reiciendis tenetur dicta sint, minus dolor aliquam quae pariatur eum. Sit, soluta, quis accusantium eaque id adipisci nesciunt veniam fugiat dignissimos 
         dolorum, temporibus neque est asperiores deserunt officiis? Sequi minus corporis porro aliquam omnis eos aut atque facilis totam quaerat enim fugiat excepturi placeat eaque voluptatibus eveniet aliquid, reiciendis, similique, distinctio veritatis assumenda? Consectetur corrupti 
         iste voluptatibus pariatur at ratione voluptate cumque vero nulla inventore nesciunt fugit ipsa, 
         voluptas officiis ex quae maiores quaerat tenetur libero rem? Fugit quod perspiciatis velit dignissimos, incidunt dolor modi cum eius maiores.</p>
        """
       

    
        

        msg.Attachments.Add(os.getcwd() +"\\pythonpdf.pdf") # attachments files are added here
        
        
        msg.Send() # mail will be send to mails included in csv file
            
print("All mail sent!")

