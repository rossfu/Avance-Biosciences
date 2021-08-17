#Calculations Performed on Form10 QPCR Data
import pandas as pd


#Extract Sample Names

def Get_Sample_Names(form10):
    Sample_Name_List = []
    for x in range(1,len(form10)-1):
        if (form10['Sample Name'][x]) == (form10['Sample Name'][x-1]):
            Sample_Name_List.append(form10['Sample Name'][x])
    return Sample_Name_List



#Extract Quantity Means
def QuantityMean(form10):
    
    Quantity_Means = {}
    Quantity_Mean_List = []
    
    for x in range(1,len(form10)-1):
        if str(form10['Sample Name'][x]).startswith('ST') == True:
            break
        if (form10['Sample Name'][x]) == (form10['Sample Name'][x-1]):
            Quantity_Mean_List.append(form10['Qty Mean'][x])
            
            mean = ((form10['Quantity'][x])+(form10.Quantity[x-1]))/ 2
            Quantity_Means[form10['Sample Name'][x]] = mean
    
    return Quantity_Means


#Quantity Means DataFrame
def QtyMean_Dataframe(form10):
    
    Qty_Means = pd.DataFrame()
    QtyMean_List = []
    Sample_names_list = []
    
    for x in range(1,len(form10)-1):
        
        if (form10['Sample Name'][x]) == (form10['Sample Name'][x-1]) and str(form10['Sample Name'][x]).startswith('ST') == False:
            QtyMean_List.append((form10['Quantity'][x])+(form10.Quantity[x-1])/ 2)
            Sample_names_list.append(form10['Sample Name'][x])

        if (form10['Sample Name'][x]) == 'NTC' and (form10['Sample Name'][x-1]) == 'NTC':
            Qty_Means['NTC'] = (form10['Quantity'][x]+form10['Quantity'][x-1])/2
        
    Qty_Means['Sample Names'] = Sample_names_list
    Qty_Means['Qty Mean'] = QtyMean_List
    
    return Qty_Means



#Extract SPK's Ct Information
def SPK_Ct_dataframe(form10):
    
    SPK_Cts = pd.DataFrame()
    SPK_names = []
    SPK_values = []
    SPK_value_differences = []
    Inhibition_true_false = []
    SPK_control_values = []
    
    #SPK Control Average
    for x in range(len(form10)):
        if form10['Sample Name'][x] == 'SPK CONTROL':
            if form10['Ct'][x] == 'Undetermined':
                SPK_avg = 'Undetermined'
            else:
                SPK_control_values = list(SPK_control_values)
                SPK_control_values.append(form10['Ct'][x])
                SPK_control_values = pd.to_numeric(SPK_control_values)
                SPK_avg = sum(SPK_control_values) / len(SPK_control_values)
    
    #Extracting +SPK Information
    for x in range(0,len(form10)):
        
        if '+SPK' in str(form10['Sample Name'][x]):
            SPK_names.append(form10['Sample Name'][x])
            SPK_values.append(form10['Ct'][x])

            if SPK_avg == 'Undetermined':
                SPK_Cts['SPK Value Differences'] = 'Undetermined'
                Inhibition_true_false.append('True')
                
            else:
                SPK_value_difference_calculation = abs(SPK_avg - float(form10['Ct'][x]))
                SPK_value_differences.append(SPK_value_difference_calculation)
            
                if SPK_value_difference_calculation >= 1:
                    Inhibition_true_false.append('True')   
                else:
                    Inhibition_true_false.append('False')

            
                
    #Displaying Results in DataFrame
    SPK_Cts['Sample Names'] = SPK_names
    SPK_Cts['SPK Values'] = SPK_values
    SPK_Cts['SPK Control Value'] = SPK_avg
    SPK_Cts['Inhibition?'] = Inhibition_true_false

    if SPK_avg != 'Undetermined':
        SPK_Cts['SPK Value Differences'] = SPK_value_differences

    
    return SPK_Cts



#Comparing Cts of Replicates and testing LOQ requirement
def Replicate_Diff(form10):
    
    Replicate_Diff_Dict = {}
    Ct_diff = 0
    
    for x in range(1,len(form10)):
        
        if (form10['Sample Name'][x]) == (form10['Sample Name'][x-1]):
            
            #Diff type and above LOQ
            if form10['Ct'][x].isalpha() != form10['Ct'][x-1].isalpha() and (form10['Quantity'][x] + form10['Quantity'][x-1] >= 25):
                
                Ct_diff = 'PROBLEM: 1 Undetermined & 1 Ct and Quantity > 25' #if 1 is above 25
            
            #Diff type and below LOQ
            elif form10['Ct'][x].isalpha() != form10['Ct'][x-1].isalpha() and (form10['Quantity'][x] + form10['Quantity'][x-1] < 25):
                
                Ct_diff = '1 Undetermined and 1 Ct but Quantity < 25'
                
            #2 undetermined
            elif form10['Ct'][x].isalpha() == True and form10['Ct'][x-1].isalpha() == True:
                Ct_diff = 'Undetermined'
            
            #Large diff and above LOQ
            elif abs(float(form10['Ct'][x]) - float(form10['Ct'][x-1])) > 1 and form10['Qty Mean'][x] > 25:
                Ct_diff = 'PROBLEM: LARGE DIFFERENCE'
                                             
            else:
                Ct_diff = float(form10['Ct'][x]) - float(form10['Ct'][x-1])
        
    Replicate_Diff_Dict[form10['Sample Name'][x]] = Ct_diff
                                        
    return Replicate_Diff_Dict


def Replicate_Diff_Dataframe(form10):
    
    Replicate_Diff_Dataframe = pd.DataFrame()
    
    Ct_diff_list = []
    Sample_list = []
    
    for x in range(1,len(form10)):
        
        if (form10['Sample Name'][x]) == (form10['Sample Name'][x-1]):

            Sample_list.append(form10['Sample Name'][x])
            
            #Diff type and above LOQ
            if form10['Ct'][x].isalpha() != form10['Ct'][x-1].isalpha() and (form10['Quantity'][x] + form10['Quantity'][x-1] > 25):
                
                Ct_diff_list.append('PROBLEM: 1 Undetermined & 1 Ct and Quantity > 25')
                
            #Diff type and below LOQ
            elif form10['Ct'][x].isalpha() != form10['Ct'][x-1].isalpha() and (form10['Quantity'][x] + form10['Quantity'][x-1] <= 25):
                
                Ct_diff_list.append('1 Undetermined and 1 Ct but Quantity < 25')
                
            #2 undetermined
            elif form10['Ct'][x].isalpha() == True and form10['Ct'][x-1].isalpha() == True:
                Ct_diff_list.append('Undetermined')
            
            #Large diff and above LOQ
            elif abs(float(form10['Ct'][x]) - float(form10['Ct'][x-1])) > 1 and form10['Qty Mean'][x] > 25:
                Ct_diff_list.append('PROBLEM: LARGE DIFFERENCE')
                                             
            else:
                Ct_diff_list.append(float(form10['Ct'][x]) - float(form10['Ct'][x-1]))
        
    Replicate_Diff_Dataframe['Sample Name'] = Sample_list
    Replicate_Diff_Dataframe['Ct diff'] = Ct_diff_list
                                        
    return Replicate_Diff_Dataframe
