import matplotlib.pyplot as plt
import xlrd
import numpy as np
def smail(x,y,z):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    f='teatimecomedy@gmail.com'
    t=y
    s=z
    msg=MIMEMultipart()
    msg['from']=f
    msg['to']=t
    msg['subject']=s
    body='This is the report of Student of'
    msg.attach(MIMEText(body,'plain'))
    filename='Progress_report.jpg'
    attachment=open(x,'rb')
    p=MIMEBase('application','octet-stream')
    p.set_payload((attachment).read())
    encoders.encode_base64(p)
    p.add_header('Content-Disposition','attachment;filename=%s'%filename)
    msg.attach(p)
    text=msg.as_string()
    server=smtplib.SMTP('smtp.gmail.com',587)
    server.starttls()
    server.login(f,'saai saai')
    server.sendmail(f,t,text)
    server.close()

print('Welcome to Marks Management System')
wb=xlrd.open_workbook('/home/rajkumar/Desktop/machine_learning_training/Marks.xlsx')
sheet1=wb.sheet_by_index(0)
sheet2=wb.sheet_by_index(1)
f='teatimecomedy@gmail.com'
while(1):
    print('\n1.Student Wise Results\n2.Subject Wise Results')
    print('3.Subject Wise Comparision\n4.Send Reports to all Parents\n5.Send Reports to Parents of weak students\n6.Send Subject wise reports to Respective Subject Teachers\n7.Exit')
    s=int(input('Select any thing from above Menu:\n'))
    if(s==1):
        snames=[sheet1.cell_value(k,0) for k in range(1,sheet1.nrows)]
        print('Student Wise Results\n')
        while(1):
            sname=input('Enter the Name of the Student')
            if sname in snames:
                m=snames.index(sname)
                m=m+1
                x1=[2,4,6,8,10,12]
                y1=[sheet1.cell_value(m,i) for i in range(2,7)]
                y1.append(sheet1.cell_value(m,8))
                y2=[35,35,35,35,35,35]
                y3=[60,60,60,60,60,60]
                labels=['Maths','Physics','Chemistry','Biology','English','Overall Percentage']
                plt.bar(x1,y1,tick_label=labels,color='green',width=0.8)
                plt.plot(x1,y2,linestyle='dashed',label='Fail',color='red')
                plt.plot(x1,y3,linestyle='dashed',color='orange',label='weak')
                plt.ylim(0,100)
                plt.xlabel('Subjects')
                plt.ylabel('Marks')
                plt.title('Results of %s'%sname)
                for j in range(6):
                    plt.text(x1[j],y1[j],'%.1f'%y1[j])
                plt.legend()
                plt.show()
                plt.close()
                break
            else:
                print('Check the name!Try Again!')
    elif(s==2):
        print('Subject Wise Results\n')
        subnames=[sheet1.cell_value(0,k) for k in range(2,7)]
        while(1):
            subname=input('Enter the Name of subject You want:')
            if subname in subnames:
                m=subnames.index(subname)
                x1=[]
                for i in range((sheet1.nrows)-1):
                   print(i+1)
                   x1.append(i+1)
                y1=[sheet1.cell_value(j,m+2) for j in range(1,sheet1.nrows)]
                y2=[35,35,35,35,35,35]
                y3=[60,60,60,60,60,60]
                labels=[sheet1.cell_value(w,0) for w in range(1,sheet1.nrows)]
                plt.bar(x1,y1,tick_label=labels,color='green',width=0.8)
                plt.plot(x1,y2,linestyle='dashed',label='Fail',color='red')
                plt.plot(x1,y3,linestyle='dashed',color='orange',label='weak')
                plt.ylim(0,100)
                plt.xlabel('Student Name')
                plt.ylabel('Marks')
                plt.title('Results of %s'%subname)
                for j in range((sheet1.nrows)-1):
                    plt.text(x1[j],y1[j],'%.1f'%y1[j])
                plt.legend()
                plt.show()
                plt.close()
                break
            else:
                print('Check the subject name and try again!')
    elif(s==3):
        print('Overall Subject Wise Comparision\n')
        x1=[2,4,6,8,10]
        y1=[]
        for j in range(5):
            a=[sheet1.cell_value(i+1,j+2) for i in range((sheet1.nrows)-1)]
            d=np.array(a)
            e=np.mean(d)
            y1.append(e)
        y2=[35,35,35,35,35]
        y3=[60,60,60,60,60]
        labels=['Maths','Physics','Chemistry','Biology','English']
        plt.bar(x1,y1,tick_label=labels,color='green',width=0.8)
        plt.plot(x1,y2,linestyle='dashed',label='Fail',color='red')
        plt.plot(x1,y3,linestyle='dashed',color='orange',label='weak')
        plt.ylim(0,100)
        plt.xlabel('Subjects')
        plt.ylabel('Marks')
        plt.title('Subject wise comparision')
        for j in range(5):
            plt.text(x1[j],y1[j],'%.1f'%y1[j])
        plt.legend()
        plt.show()
        plt.close()
    elif(s==4):
        print('Send Reports to all Parents')
        snames=[sheet1.cell_value(k,0) for k in range(1,sheet1.nrows)]
        for sname in snames:
            m=snames.index(sname)
            m=m+1
            x1=[2,4,6,8,10,12]
            y1=[sheet1.cell_value(m,i) for i in range(2,7)]
            y1.append(sheet1.cell_value(m,8))
            y2=[35,35,35,35,35,35]
            y3=[60,60,60,60,60,60]
            labels=['Maths','Physics','Chemistry','Biology','English','Overall Percentage']
            plt.bar(x1,y1,tick_label=labels,color='green',width=0.8)
            plt.plot(x1,y2,linestyle='dashed',label='Fail',color='red')
            plt.plot(x1,y3,linestyle='dashed',color='orange',label='weak')
            plt.ylim(0,100)
            plt.xlabel('Subjects')
            plt.ylabel('Marks')
            plt.title('Results of %s'%sname)
            for j in range(6):
                plt.text(x1[j],y1[j],'%.1f'%y1[j])
            plt.legend()
            plt.savefig('graph.png')
            smail('graph.png',sheet1.cell_value(m,1),'Techienest Result Reports of your Son/Daughter')
            plt.close()
            print('sent')
        print('All Mails sent\n')
    elif(s==5):
        print('Send Reports to Parents of Weak Students')
        snames=[sheet1.cell_value(k,0) for k in range(1,sheet1.nrows)]
        for sname in snames:
            m=snames.index(sname)
            m=m+1
            if sheet1.cell_value(m,8)<60:
                x1=[2,4,6,8,10,12]
                y1=[sheet1.cell_value(m,i) for i in range(2,7)]
                y1.append(sheet1.cell_value(m,8))
                y2=[35,35,35,35,35,35]
                y3=[60,60,60,60,60,60]
                labels=['Maths','Physics','Chemistry','Biology','English','Overall Percentage']
                plt.bar(x1,y1,tick_label=labels,color='green',width=0.8)
                plt.plot(x1,y2,linestyle='dashed',label='Fail',color='red')
                plt.plot(x1,y3,linestyle='dashed',color='orange',label='weak')
                plt.ylim(0,100)
                plt.xlabel('Subjects')
                plt.ylabel('Marks')
                plt.title('Results of %s'%sname)
                for j in range(6):
                    plt.text(x1[j],y1[j],'%.1f'%y1[j])
                plt.legend()
                plt.savefig('graph.png')
                smail('graph.png',sheet1.cell_value(m,1),'Weak Result Reports of your Son/Daughter')
                plt.close()
                print('sent')
        print('All Mails sent\n')
    elif(s==6):
        print('Send Subject wise reports to respective Subject teachers\n')
        subnames=[sheet1.cell_value(0,k) for k in range(2,7)]
        for subname in subnames:
            m=subnames.index(subname)
            x1=[]
            for i in range((sheet1.nrows)-1):
               x1.append(i+1)
            y1=[sheet1.cell_value(j,m+2) for j in range(1,sheet1.nrows)]
            y2=[35,35,35,35,35,35]
            y3=[60,60,60,60,60,60]
            labels=[sheet1.cell_value(w,0) for w in range(1,sheet1.nrows)]
            plt.bar(x1,y1,tick_label=labels,color='green',width=0.8)
            plt.plot(x1,y2,linestyle='dashed',label='Fail',color='red')
            plt.plot(x1,y3,linestyle='dashed',color='orange',label='weak')
            plt.ylim(0,100)
            plt.xlabel('Student Name')
            plt.ylabel('Marks')
            plt.title('Results of %s'%subname)
            for j in range((sheet1.nrows)-1):
                plt.text(x1[j],y1[j],'%.1f'%y1[j])
            plt.legend()
            plt.savefig('graph.png')
            smail('graph.png',sheet2.cell_value(m+1,2),'Your Subject Analysis of students')
            print('sent')
            plt.close()
        print('allmails sent')
    elif(s==7):
        break
    else:
        print('Invalid Entry!Try Again')
