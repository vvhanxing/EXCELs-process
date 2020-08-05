import pandas as pd
import os
from  collections import OrderedDict


def get_excel_block(file_name):
    
    sheet = pd.read_excel(file_name, sheet_name=0)
    nrows =  sheet.shape[0]
    ncols = sheet.columns.size
    if ncols ==0:
        

        sheet = pd.read_excel(file_name, sheet_name=1)
        nrows =  sheet.shape[0]
        ncols = sheet.columns.size
        print("next sheet find")
        if  ncols ==0:
            sheet = pd.read_excel(file_name, sheet_name=2)
            nrows =  sheet.shape[0]
            ncols = sheet.columns.size
            print("next next sheet find")
        
    column_names = sheet.columns
    
    #print(column_names)
    data_dir = {}
    #print(file_name)
    #print(column_names)
    
    try :
        for r in range(nrows):
            for i in range(ncols):
                data=sheet.iloc[r,i]
            #print(column_names[i],data)
                data_dir[column_names[i]] = data
            yield data_dir
    except IndexError:
        print("nothing")
        yield {}



def get_excel_category(file_name):

    sheet = pd.read_excel(file_name, sheet_name=0)
    nrows =  sheet.shape[0]
    ncols = sheet.columns.size
    if ncols ==0:
        
        sheet = pd.read_excel(file_name, sheet_name=1)
        nrows =  sheet.shape[0]
        ncols = sheet.columns.size
        print("next sheet find")
        if  ncols ==0:
            sheet = pd.read_excel(file_name, sheet_name=2)
            nrows =  sheet.shape[0]
            ncols = sheet.columns.size
            print("next next sheet find")
    column_names = sheet.columns
    return list(column_names)





#input(files_list)

def get_all_info(files_list):


    all_category = []  
     
    for i,name in enumerate(files_list):
        new_category =  get_excel_category(name)
        if len(new_category)!=0:
            all_category.extend(new_category)
        else:
            print(name)
            


    for i in  all_category:
        print("-> ",i)


    all_category_dict = { x: [] for x in all_category}

    all_category_dict["file_source"] = []
   
    for i,name in enumerate(files_list ):

        for  block_dict in get_excel_block(name):
        
            for all_category_key in all_category_dict.keys():
                if all_category_key!="file_source":
                
                    all_category_dict[all_category_key].append(block_dict.get(all_category_key,"\\"))
                else:
                    all_category_dict["file_source"].append(name)
    return all_category_dict
       
            


def creat_all_data_xlsx(files_list):
    all_category_dict =  get_all_info(files_list)

    df=pd.DataFrame(all_category_dict)#构造原始数据文件

    df.to_excel('all_data.xlsx')#生成Excel文件，并存到指定文件路径下

    print("finish")




def select_info(ID,category,all_data='all_data.xlsx'):#在总表中依据category的ID搜索,并将该行信息建立列表，列表中每对信息为字典
    sheet = pd.read_excel(all_data, sheet_name=0)
    #print(sheet)
    nrows =  sheet.shape[0]
    ncols = sheet.columns.size
    column_names = sheet.columns

    
    data_list = []
    
    indexs = []
    try :
        for index ,i in enumerate(sheet[category]):
            if i==ID:
                indexs.append( index)
                
        for index in indexs:
            data_dir = {}
            for i in range(ncols):
            
                data=sheet.iloc[index,i]
                #if data_dir.get(column_names[i],"Nothing") !="Nothing":
                data_dir[column_names[i]] = data
                #else:
                    #print("repet",data)
            data_list.append(data_dir)

    except IndexError:
        print("nothing")    
    
    return data_list
    

def get_standard(string):
    if "\n" in string:
        string = string[:-1]

    string=" ".join([x for x in string.split(" ")if x!=""])
    
    if string=="DEL No":
        string = 'DEL No.'
    if string == "type":
        string ="Type"

        
    return string






def get_category_names(file_name):
    sheet = pd.read_excel(file_name, sheet_name=0)
    nrows =  sheet.shape[0]
    ncols = sheet.columns.size
    if ncols ==0:
        

        sheet = pd.read_excel(file_name, sheet_name=1)
        nrows =  sheet.shape[0]
        ncols = sheet.columns.size
        print("next sheet ")
        if  ncols ==0:
            sheet = pd.read_excel(file_name, sheet_name=2)
            nrows =  sheet.shape[0]
            ncols = sheet.columns.size
            print("next next sheet ")
        
    column_names = sheet.columns

    return [ get_standard(x)  for x in list(column_names)]






def ouput(Target_ID,Campaign_ID,Nsynthon_ID,model_excel_name):
    #寻找Target_ID,Campaign_ID,Nsynthon_ID含有的各个excel数据，并返回带顺序的字典


    cat_data_dir ={}#在总表条件寻找信息并汇总到字典,由于可以找全，优先级别最高。此后信息分为补充信息和冲突信息
    file_source=[]
    c = 0
    for block_dir in select_info( Nsynthon_ID,"NsynthonID"):
        if  block_dir.get("NsynthonID","\\")==Nsynthon_ID and block_dir.get("Target ID","\\")==Target_ID and block_dir.get("Campaign ID","\\")==Campaign_ID:
            c+=1
            #print(c,">>",block_dir)
            for key,value in block_dir.items():
                if value!='\\':
                    if key=="file_source":
                        file_source.append(value)
                    else:   
                        cat_data_dir[key]=[value]
                        #print(key,value)
            cat_data_dir["file_source"] = " ".join(file_source)




    cat_data_dir_0 ={}  #在总表条件寻找信息并汇总到字典
    #file_source=[]
    for block_dir in select_info( Target_ID,"Target ID"):
        if  block_dir.get("NsynthonID","\\")=="\\" and block_dir.get("Target ID","\\")==Target_ID and block_dir.get("Campaign ID","\\")==Campaign_ID:
            for key,value in block_dir.items():
                if value!='\\':
                    if key=="file_source":
                        file_source.append(value)
                    else:   
                        cat_data_dir_0[key]=[value]
            cat_data_dir_0["file_source"] = " ".join(file_source)    

    


    cat_data_dir_1 ={}  #在总表条件寻找信息并汇总到字典
    
    for block_dir in select_info( Nsynthon_ID,"NsynthonID"):
        #print("here")
        if  block_dir.get("NsynthonID","\\")==Nsynthon_ID and block_dir.get("Target ID","\\")=="\\"and block_dir.get("Campaign ID","\\")=="\\":
            for key,value in block_dir.items():
                if value!='\\':
                    if key=="file_source":
                        file_source.append(value)
                    else:   
                        cat_data_dir_1[key]=[value]
                        #print("------------------------------key",key,value)
            cat_data_dir_1["file_source"] = " ".join(file_source)

                 
                
    cat_data_dir = {**cat_data_dir_0,**cat_data_dir} #之前所有汇总的整合一起
    cat_data_dir = {**cat_data_dir_1,**cat_data_dir} #之前所有汇总的整合一起
    
    #print("finish")
    #input()     
    #for key,value in cat_data_dir.items():

            
        #print('{:^30}  {:<30}'.format(key,value))

        
    
    Ordered_category_names = get_category_names(  model_excel_name)
    print(Ordered_category_names) #建立有序空字典，顺序是模板的标题名
    Ordered_category_names = [   get_standard(n) for n in Ordered_category_names] #把model中标题标准化放入新表中
    #Ordered_category_names.append("file_source")
    Ordered_list = list(zip(*[Ordered_category_names,[["\\"] for x in range(len(Ordered_category_names  ))]]))
    cat_data_dir_order = OrderedDict( Ordered_list)

    

    # cat_data_dir ,Ordered_category_names,cat_data_dir_order


    
    for key,value in cat_data_dir.items():#把汇总的信息按照模板顺序放到新表里
        if get_standard(key) in Ordered_category_names: #把cat_data_dir标题标准化放入新表中
            #print(key,value)
            
            cat_data_dir_order[get_standard(key)] = value


    return cat_data_dir_order




def save_output_excel(input_excel_name,model_excel_name,save_name):

    sheet = pd.read_excel(input_excel_name, sheet_name=0)
    
    nrows =  sheet.shape[0]
    ncols = sheet.columns.size
    if ncols ==0:
        
        sheet = pd.read_excel(file_name, sheet_name=1)
        nrows =  sheet.shape[0]
        ncols = sheet.columns.size
        print("next sheet ")
        if  ncols ==0:
            sheet = pd.read_excel(file_name, sheet_name=2)
            nrows =  sheet.shape[0]
            ncols = sheet.columns.size
            print("next next sheet ")
    column_names = sheet.columns
    
    print(sheet)
    
    
    #sheet_ = pd.read_excel( model_excel_name, sheet_name=2)
    #print(sheet_)
    #input("finish")
    Target_ID,Campaign_ID ,Nsynthon_ID = sheet["Target ID"][0],sheet["Campaign ID"][0],sheet["NsynthonID"][0]
    cat_data_dir_order = ouput(Target_ID,Campaign_ID,Nsynthon_ID,model_excel_name)
    
    print(cat_data_dir_order)    
    for i in range(nrows):
        if i>0:
        
            print(sheet["Target ID"][i],sheet["Campaign ID"][i],sheet["NsynthonID"][i])
   
            Target_ID,Campaign_ID ,Nsynthon_ID  = sheet["Target ID"][i],sheet["Campaign ID"][i],sheet["NsynthonID"][i]

            

            cat_data_dir_order_add = ouput(Target_ID,Campaign_ID,Nsynthon_ID,model_excel_name)
        
    


        
            for key,value in cat_data_dir_order_add.items():  



                print(key,value)

                cat_data_dir_order[key].extend(value)
               


        
    #cat_data_dir_order["Entry"].append('\\')
    print(cat_data_dir_order)
    dataframe = pd.DataFrame(cat_data_dir_order)
    
    dataframe.to_excel(save_name)
    




    
            
    print("------------------------------------------------------")





def for_panl_data(input_file_names):
    
    exceldata = []
    for file_name in input_file_names:
    
        #sheet_dict = read_excel(file_name)
        titles = get_category_names(file_name)

        for i,titel in enumerate(titles):
        #print(titel,sheet_dict[titel])
            exceldata.append([i,i,titel,file_name])
    return exceldata



if __name__=="__main__":
    

    #folder  = "7-28//"   
    #files_list = [folder+x for x in os.listdir(folder) if ".xlsx" in x]
    #creat_all_data_xlsx(files_list)
    



    model_excel_name = "输出/输出-Custom compound delivery tracking list for Client- PM.xlsx"
    input_excel_name = 'input.xlsx'
    save_name = '6.xlsx'
    save_output_excel(input_excel_name,model_excel_name,save_name)








