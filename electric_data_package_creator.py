from tkinter import *
import pyodbc
import pandas as pd
 
#Connect to SQL Server -------------------------------------------------------
server = 'xxxxxxx'                          #server name
user = 'xxxxxxx'                                  #username
pword = 'xxxxxxx'                                    #password
#-----------------------------------------------------------------------------
 
 
#===================   START GUI   ===========================================
def testphase_select():
    global selected_testphases, phase_list
    selected_testphases = []
    for i in phase_list.curselection():
        selected_testphases.append(phase_list.get(i))
    root1.destroy()
 
 
def slicenum_submit():                                                         #slice number select function
    global selected_slcs, slc_list, testphases, query_extension, query, phase_list
    selected_slcs = []
    testphases = []
    for i in slc_list.curselection():
        selected_slcs.append(slc_list.get(i))
   
    query_extension = ''
   
    for i in list(range(1, len(selected_slcs))):
        query_extension += " OR tbl_dutInfo.dutSubA_SN='"+selected_slcs[i]+"'"
       
    query = "SELECT DISTINCT test_Phase FROM tbl_dutInfo, tbl_vector WHERE tbl_dutInfo.dutSubA_SN='"+selected_slcs[0]+\
    "'"+query_extension+" ORDER BY test_Phase"
    cursor.execute(query)
   
    for i in cursor:
        testphases.append(str(i)[2:-4])
 
    phase_list = Listbox(root1, selectmode =  'multiple')
    phase_list.grid(row=4, column=1)
    
    for testphase in testphases:
        phase_list.insert(END, testphase)
   
    
def partnum_submit():                                                          #part number submit function
    global slc_list
    slc_list = Listbox(root1, selectmode='multiple')
    slc_list.grid(row=3,column=1)
    for slc in slcs:
        slc_list.insert(END,slc)
 
def partnum_click(partnum):                                                    #part number select function
    global pn, slcs
    pn = partnum
    slcs = []
   
    cursor.execute("SELECT DISTINCT dutSubA_SN FROM tbl_dutInfo WHERE tbl_dutInfo.dutTopA_PN='"+pn+"'")
    for i in cursor:
        slcs.append(str(i)[2:-4])
 
def program_click():                                                           #program select function
    global prog, cnxn, cursor, pns
    prog = str(program_entry.get())
    pns = []
 
    cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+prog+';UID='+user+';PWD='+ pword)
    cursor = cnxn.cursor()
    cursor.execute("USE " + prog) #ensures connection to the database of interest
    cursor.execute("SELECT DISTINCT dutTopA_PN FROM tbl_dutInfo")
   
    for i in cursor: 
        pns.append(str(i)[2:-4])
 
    clicked = StringVar()
    clicked.set('pick a unit')
     
    partnum_menu = OptionMenu(root1, clicked, *pns, command = partnum_click)   
    partnum_menu.grid(row=2, column=1)   
         
        
root1 = Tk()
root1.title('Electronic Test Data Retrieval for EIDP')
 
program_label  = Label(root1, text = 'Program: ').grid(row=1, column=0, sticky=W)
program_entry  = Entry(root1)
program_entry.grid(row=1, column=1, sticky=W)
program_button = Button(root1, text = 'Submit', command = program_click).grid(row=1, column=2, sticky=W)
 
partnum_label  = Label(root1, text = 'Part Number: ').grid(row=2, column=0, sticky=W)
partnum_button = Button(root1, text = 'Submit', command = partnum_submit).grid(row=2,column=2,sticky=W)
 
slicenum_label = Label(root1, text = 'Slice Serial Numbers: ').grid(row=3, column=0, sticky=W)
slicenum_button = Button(root1, text = 'Submit', command = slicenum_submit).grid(row=3, column=2, sticky=W)
 
testphase_label = Label(root1, text = 'Test Phases: ').grid(row=4, column=0, sticky=W)
testphase_button = Button(root1, text = 'Submit', command=testphase_select).grid(row=4,column=2, sticky=W)
 
root1.mainloop()
#-----------------------------------------------------------------------------
#====================   END GUI   ============================================
###############################################################################
###############################################################################
 
 
#Initial Performance Section
for slc in selected_slcs:
    for phase in selected_testphases:
        
        i=0
        cursor.execute("SELECT  test_Phase, test_Parameter, test_Desc, ptNum, x, y "+
                      
                       "FROM tbl_dutInfo, tbl_vector, tbl_vectorData "+
                      
                       "WHERE tbl_dutInfo.dutTopA_PN = '"+pn+"' AND tbl_dutInfo.dutInfo_ID = tbl_vector.dutInfo_ID "+
                       "AND tbl_vector.archive = 1 AND tbl_vector.vector_ID = tbl_vectorData.vector_ID AND dutSubA_SN='"+slc+\
                       "' AND test_Phase = '"+phase+"'")  #potentially add ORDER BY test_Phase, test_Parameter, test_Desc, ptNum
                
        #gets the column names from the sql query
        Desc = cursor.description
        Column_names = [x[0] for x in Desc]
       
        #makes dataframe with column names
        raw_data = pd.DataFrame(columns = Column_names)
       
        i = 0 #row counter
        Num_col = list(range(0,len(Column_names)))
   
        for row in cursor: 
            raw_data.loc[i,:] = row
            i+=1
 
        w = pd.ExcelWriter('SN'+slc+'_'+phase+'.xlsx')
 
        #creates each excel file
        parameters = raw_data.test_Parameter.drop_duplicates().to_list()
   
        #creates each sheet
        for parameter in parameters:
            parameter_df = raw_data[raw_data.test_Parameter == parameter]        
            
            sheet_name = phase+' '+parameter       
            sheet_df = pd.DataFrame()
           
            #col_sets = list(range(0,len(parameter_df.test_Desc.drop_duplicates())))
            descs = parameter_df.test_Desc.drop_duplicates().to_list()
   
            #creates each column set
            for desc in descs:
                desc_df = parameter_df[parameter_df.test_Desc == desc]
               
                headers = pd.DataFrame([[desc, '', '']], columns = ['x','y',0])               
                desc_df = pd.concat([headers, desc_df], axis = 0)
                desc_df.reset_index(drop = True, inplace = True)   
                blank = pd.DataFrame(data = ['']*len(desc_df))
               
                sheet_df = pd.concat([sheet_df, desc_df[['x','y']], blank], axis = 1)
   
            sheet_df.to_excel(w, sheet_name = sheet_name, index = False, header = False)  #slc \ w bc slc is the folder
           
        print(slc, phase)
        w.save()
        w.close()
           
        
cursor.close()   #close the cursor
cnxn.close()    #close the connection
