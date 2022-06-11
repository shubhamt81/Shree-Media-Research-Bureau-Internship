
def changed():
    msg='Caste of Surname is changed Successfully!\n'+sirname.get()+'->'+caste.get()
    
    if(sirname.get()=="Select Sirname" or caste.get()=="Select Caste" or i.get()=="Select Row for sirname" or j.get()=="Select Row for caste"):
        messagebox.showerror("ALERT","SELECT ALL THE VALUES")
        
    else:
        
        r=int(i.get())
        p=int(j.get())
        print(p,r,sirname.get(),caste.get())
        for k in range(3,100):
            cell_obj = sheet_obj.cell(row = k, column =r)
            print(cell_obj.value)
            if(cell_obj.value==None):
                continue
            else:
                lst=list(map(str,cell_obj.value.split()))
                print(lst[-1])
                if(lst[-1]==sirname.get()):
                    print("FOUND",p,r,k)
                    sheet_obj.cell(row=k,column=p).value=caste.get()
        wb_obj.save(path)
        messagebox.showerror("ALERT",msg)
   