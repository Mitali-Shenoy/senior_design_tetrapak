
# %%
import random
from re import T
import pandas as pd
import numpy as np
import tkinter as tk 
import datetime
import os
#from xlsxwriter import Workbook

from tkinter import filedialog
from tkinter import *

# %%

file1 = open('output1.txt', 'w')
file2 = open('output2.txt', 'w')
# Defining data structures for doubly linked lists 

class Node: 
    def __init__(self, orderID, noRolls, noLanes, qsv, waste, no_coprint, restriction, due_date): 
        
        self.orderID = orderID
        
        if noRolls > 7:
            self.isLarge = True #flags if a single large order
        else:
            self.isLarge = False
        if orderID[-1] == 'C':  #flags if order is a Coprint job
            self.isCoprint = True
        else:
            self.isCoprint = False
        
        self.noRolls = noRolls
        self.noLanes = noLanes
        self.qsv = qsv
        self.waste = waste
        self.no_coprint = no_coprint
        self.restriction = restriction
        self.due_date = due_date

        self.start_time = None
        self.end_time = None
        self.next = None 
        self.prev = None 

class doublyLinkedList: 
    def __init__(self): 
        self.head = None

    def insertAtBeg(self,new_node): 

        # new_node = Node(orderID, noRolls, noLanes)
        new_node.next = self.head
        new_node.prev = None 


        if self.head is not None: 
            self.head.prev = new_node

        self.head = new_node
        new_node.start_time = 0
        new_node.end_time = 18 * new_node.noRolls
        last = self.head
        while (last.next is not None):
            last = last.next
            last.start_time = last.prev.end_time
            last.end_time = last.start_time + (18 * last.noRolls)


    def insertAtPos(self, prev_node, new_node):
        if prev_node is None: 
            print("This node does not exist")
            return 

        # new_node = Node(orderID, noRolls, noLanes) 
        new_node.next = prev_node.next
        prev_node.next = new_node
        new_node.prev = prev_node

        if new_node.next is not None: 
            new_node.next.prev = new_node
        
        new_node.start_time = new_node.prev.end_time
        new_node.end_time = new_node.start_time + (18 * new_node.noRolls)

        rest = new_node
        while (rest.next is not None):
            rest = rest.next
            rest.start_time = rest.prev.end_time
            rest.end_time = rest.start_time + (18 * rest.noRolls)
            

    def insertAtEnd(self, new_node): 
        # new_node = Node(orderID, noRolls, noLanes)
        last = self.head

        new_node.next = None

        if self.head is None: 
            new_node.prev = None 
            new_node.start_time = 0
            new_node.end_time = 18 * new_node.noRolls            
            self.head = new_node
            return 

        while (last.next is not None): 
            last = last.next
        
        last.next = new_node 
        new_node.prev = last
        if new_node.start_time is None:
            new_node.start_time = last.end_time
            new_node.end_time = new_node.start_time + (18 * new_node.noRolls)
        else:
            new_node.start_time = 40 + last.end_time
            new_node.end_time = new_node.start_time + (18 * new_node.noRolls)


    #def get():
    
    def printList(self, node):
 
#        print("\nTraversal in forward direction")
        text = ""
        while node:
            
            # print(" {} {} {} {} {} {} {} {} {}".format(node.orderID, node.qsv, node.noLanes, node.noRolls, node.waste, node.no_coprint, node.restriction, node.start_time, node.end_time))
            text = text + "{},{},{},{},{},{},{},{}\n".format(node.orderID, node.qsv, node.noLanes, node.noRolls, node.waste, node.no_coprint, node.restriction, node.due_date)
            # last = node
            node = node.next

        return text
    
    # This function returns size of linked list
    def findSize(self, node):
    
        res = 0
        while (node != None):
            res = res + 1
            node = node.next
        
        return res

    def CopyList(self):#Out put new Linked List that is a copy of current Linked List with out altering it. 
        # create new LinkedList
        newLinkedList = doublyLinkedList()
        current = self.head
        #below is from stackoverflow : https://stackoverflow.com/questions/36491307/how-to-copy-linked-list-in-python
        while current is not None:
            newNode = Node(current.orderID, current.noRolls, current.noLanes, current.qsv, current.waste, current.no_coprint, current.restriction, current.due_date)
            newLinkedList.insertAtEnd(newNode)
            current = current.next
        return newLinkedList
    
    def swap(self, old, new):
        
        tail = self.head
        
        while (tail.next is not None):
            tail = tail.next
       
        if old == self.head:
            new.next = self.head.next
            new.prev = None
            new.start_time = 0
            new.end_time = new.start_time + (18 * new.noRolls)
            self.head = new
            
        elif old == tail:
            new.next = None
            new.prev = tail.prev
            new.start_time = new.prev.end_time
            new.end_time = new.start_time + (18 * new.noRolls)
            tail.prev.next = new
            # tail = new
        else:
            new.prev = old.prev
            new.next = old.next
            new.start_time = new.prev.end_time
            new.end_time = new.start_time + (18 * new.noRolls)
            old.prev.next = new
            old.next.prev = new
            
            
        
        # new.start_time = new.prev.end_time
        # new.end_time = new.start_time + (18 * new.noRolls)
        cur = new
        while cur.next is not None:
            if cur.prev.qsv != cur.qsv:
                cur.start_time = cur.prev.end_time + 40
            cur.end_time = cur.start_time + (18 * cur.noRolls)
            cur = cur.next

        old.next = None
        old.prev = None
        return old
    


# llist = doublyLinkedList()
# new_node = Node("36", 4, 5,24,56,23,"y")
# llist.insertAtEnd(new_node)
# llist.insertAtEnd("37", 4, 5, 3)
# llist.insertAtEnd("39", 4, 5, 3)
# llist.insertAtEnd("42", 4, 5, 3)
# llist.insertAtBeg("64", 4, 5, 3)
# llist.insertAtBeg("93", 4, 5, 3)
# llist.insertAtBeg("72", 4, 5, 3)

# llist.printList(llist.head)

# %%
# file = open('output.txt', 'w')

tk.Tk().withdraw() # prevents an empty tkinter window from appearing

#asking the user to select their input file
file_path = filedialog.askopenfilename(title = "Select WIP Report")

# print(file_path)

#converting input excel file into a dataframe 
wr_df = pd.read_excel(file_path) #read Wip Report File (wp)

# print(wr_df)


# file = open('output.txt', 'w')

tk.Tk().withdraw() # prevents an empty tkinter window from appearing

#asking the user to select their input file
file_path = filedialog.askopenfilename(title = "Select Slit Plan")

# print(file_path)

# converting input excel file into a dataframe 


sp_df = pd.read_excel(file_path) #read Slit Plan file

# print(sp_df)


# cleaning up the input data file for WIP report 
wr_df = wr_df.iloc[3:-4,:]
wr_df.columns = wr_df.iloc[0]
wr_df = wr_df.drop(wr_df.index[0])
wr_df = pd.DataFrame(wr_df)
wr_df = wr_df[["Material"," Order", " QSV", " Lanes", " No Of Rolls", " Potential Waste Length", " POrder Due Date"]]
# we only care about WIP's that are in the warehouse after the lamination stage ('WIP Printing-Lamina')
wr_df = pd.DataFrame(wr_df[wr_df['Material']=='WIP Printing-Lamina'])
wr_df = wr_df.astype({" Order": str, " Lanes": int, " No Of Rolls": int}) #setting the datatype for columns
wr_df[" POrder Due Date"] = pd.to_datetime(wr_df[" POrder Due Date"])
wr_df["Restriction"] = None
# wr_df["Package Size"] = None



# print()
# print(sp_df.columns)

# cleaning up the in[ut data file for Slit Plan
sp_df = sp_df[[0, "Order number", "Package.Volume", "Package.Shape", "Lane Assignment", "Size", "Customer name"]]
sp_df = sp_df.astype({0 : int, "Order number": str, "Package.Volume": str, "Package.Shape": str, "Lane Assignment": str}) #setting the datatype for columns

for index, row in wr_df.iterrows():
    matching_id = sp_df[sp_df['Order number']==row[' Order']]
    volume = matching_id.iloc[0,2]
    shape = matching_id.iloc[0,3]
    # size = matching_id.iloc[0,5]

    if (volume == "250 ml") and (shape == "Edge"):
        wr_df.at[index, 'Restriction'] = '55'
    elif (volume == "250 ml") and (shape == "Base Leaf"):
        wr_df.at[index, 'Restriction'] = "55"
    elif (volume == "125 ml") and (shape == "Slim"): 
        wr_df.at[index, 'Restriction'] = "54"
    else: 
        wr_df.at[index, 'Restriction'] = "na"

    wr_df.at[index, ' POrder Due Date'] = row[" POrder Due Date"].date()
    # wr_df.at[index, 'Package Size'] = size

    
# print(wr_df)


# print(sp_df)

# %%

sorted_wr_df = wr_df.sort_values(by = [" POrder Due Date"])
# print(sorted_wr_df[:20])

# order_len = len(sorted_wr_df)
# bucket_1 = sorted_wr_df.iloc[0:int(order_len/3),:] 
# bucket_2 = sorted_wr_df.iloc[int((order_len/3)):int(2*(order_len/3))+1,:] 
# bucket_3 = sorted_wr_df.iloc[int((2*(order_len/3))+1):order_len+1,:] 


# # print("\nBucket 1")
# sorted_b1 = bucket_1.sort_values(by = [" QSV"])
# # print(sorted_b1)
# # print("\nBucket 2\n")
# sorted_b2 = bucket_2.sort_values(by = [" QSV"])
# # print(sorted_b2)
# # print("\nBucket 3\n")
# sorted_b3 = bucket_3.sort_values(by = [" QSV"])
# # print(sorted_b3)

sorted_wr_df = sorted_wr_df.reset_index()
curr_date = sorted_wr_df.at[0,' POrder Due Date']
buckets = [[]]
bucket_no = 0
# print(curr_date)
for index, row in sorted_wr_df.iterrows():
    if row[' POrder Due Date'] == curr_date or row[' POrder Due Date'] == curr_date+ datetime.timedelta(days=1):
        buckets[bucket_no].append(row)
        # curr_date = row[' POrder Due Date']
    else:
        bucket_no +=1
        buckets.append([])
        curr_date = row[' POrder Due Date']
        buckets[bucket_no].append(row)

grouped_df = pd.DataFrame()
for bucket in buckets: 
    tmp_df = pd.DataFrame(bucket)
    tmp_df = tmp_df.sort_values(by = [" QSV"])
    grouped_df = pd.concat([grouped_df, tmp_df])

# print(grouped_df)

# %%

schedule54 = doublyLinkedList() # linked list maintaining schedule for slitter 54
schedule55 = doublyLinkedList() # linkedlist maintaining schedule for slitter 55


ERT_54 = 0
ERT_55 = 0

df_54 = grouped_df[grouped_df['Restriction']=='54']
df_55 = grouped_df[grouped_df['Restriction']=='55']

# print(df_54)
# print(df_55)

t_rolls_54 = df_54[" No Of Rolls"].sum()
t_rolls_55 = df_55[" No Of Rolls"].sum()

# print(t_rolls_54)
# print(t_rolls_55)

t_qsv_54 = df_54[" QSV"].nunique()
t_qsv_55 = df_55[" QSV"].nunique()

# print(t_qsv_54)
# print(t_qsv_55)

ERT_54 = (t_rolls_54 *18) + (t_qsv_55 *40)
ERT_55 = (t_rolls_55 *18) + (t_qsv_55 *40)



# creating nodes from the input file 

# printing new lines
# print()
# # print()

curr_qsv_54 = ""
curr_qsv_55 = ""

def create_node(wip_row):

    orderID = wip_row[" Order"]
    noRolls = wip_row[" No Of Rolls"]
    noLanes = wip_row[" Lanes"]
    qsv = wip_row[" QSV"]
    waste = wip_row[" Potential Waste Length"]
    due_date = wip_row[" POrder Due Date"]

    if orderID[-1]=='C':
        matching_id = sp_df[sp_df['Order number']==orderID]
        lane_id = int(matching_id.iloc[0,0])
        matching_rows = sp_df[sp_df[0]==lane_id]
        no_coprint = len(matching_rows)
        
    else:
        no_coprint = 0

    matching_id = sp_df[sp_df['Order number']==orderID]
    volume = matching_id.iloc[0,2]
    shape = matching_id.iloc[0,3]

    if (volume == "250 ml") and (shape == "Edge"):
        restriction = "55"
    elif (volume == "250 ml") and (shape == "Base Leaf"):
        restriction = "55"
    elif (volume == "125 ml") and (shape == "Slim"): 
        restriction = "54"
    else: 
        restriction = "na"

#    print(orderID, noRolls, noLanes, qsv, waste, no_coprint, restriction)
    new_node = Node(orderID, noRolls, noLanes, qsv, waste, no_coprint, restriction, due_date)

    global ERT_54
    global ERT_55
    global curr_qsv_55
    global curr_qsv_54
    
    
    

    # if restriction == "54":
    #     # ERT_54 = ERT_54 + (18*noRolls)
    #     if curr_qsv_54 is not qsv and schedule54.head is not None: 
    #         ERT_54 = ERT_54 + 40
        
    
    # if restriction == "55": 
    #     # ERT_55 = ERT_55 + (18*noRolls)   
    #     if curr_qsv_55 is not qsv and schedule55.head is not None:
    #         ERT_55 = ERT_55 + 40     

    return new_node

#for index, row in sorted_b1.iterrows():
#    new_node = create_node(row)
#    schedule54.insertAtEnd(new_node)

#temp = sorted_b1.iloc[0]
#new_node = create_node(temp)
# schedule54.insertAtBeg(new_node)
# temp = sorted_b1.iloc[2]
# new_node = create_node(temp)
#temp = schedule54.CopyList()
# temp = schedule54.head
# while (temp.next is not None): 
#     temp = temp.next

# schedule54.insertAtPos(temp.prev, new_node)
# schedule54.printList(schedule54.head)

def scanToInsert(node):
    global curr_qsv_54
    global curr_qsv_55
    global ERT_55
    global ERT_54

    if schedule54.head is None: 
        if schedule55.head is None:
            if node.restriction == '54' or node.restriction =='na':
                schedule54.insertAtEnd(node)
                curr_qsv_54 = node.qsv
                return
            elif node.restriction == '55':
                schedule55.insertAtEnd(node)
                curr_qsv_55 = node.qsv
                return
            # else:
            #     val = random.randint(54, 55)
            #     if val == 54:
            #         schedule54.insertAtEnd(node)
            #         curr_qsv_54 = node.qsv
            #         return
            #     else:
            #         schedule55.insertAtEnd(node)
            #         curr_qsv_55 = node.qsv
            #         return

    cur1 = schedule54.head
    cur2 = schedule55.head
    if cur1 is not None:
        while (cur1.next is not None):
            cur1 = cur1.next
    if cur2 is not None:
        while (cur2.next is not None):
            cur2 = cur2.next

    if schedule54.head is None: 
        if schedule55.head is not None:
            if cur2.qsv == node.qsv:
                schedule55.insertAtEnd(node)
                curr_qsv_55 = node.qsv
                return
            else:
                if node.restriction == '55':
                    schedule55.insertAtEnd(node)
                    curr_qsv_55 = node.qsv
                    return
                elif node.restriction == '54':
                    schedule54.insertAtEnd(node)
                    curr_qsv_54 = node.qsv
                    return
                else:
                    if node.restriction == 'na' and ERT_54 < ERT_55:
                        if cur1 is not None:
                            if cur1.qsv is not node.qsv:
                                ERT_54 += (40 +18*node.noRolls)
                            else:
                                ERT_54 += (18*node.noRolls)
                        schedule54.insertAtEnd(node)
                        curr_qsv_54 = node.qsv
                        return
                    else:
                        if cur2 is not None:
                            if cur2.qsv is not node.qsv:
                                ERT_55 += (40 +18*node.noRolls)
                            else:
                                ERT_55 += (18*node.noRolls)
                        schedule55.insertAtEnd(node)
                        curr_qsv_55 = node.qsv
                        return

    elif schedule54.head is not None: 
        if schedule55.head is None:
            if cur1.qsv == node.qsv:
                schedule54.insertAtEnd(node)
                curr_qsv_54 = node.qsv
                return
            else:
                if node.restriction == '54':
                    schedule54.insertAtEnd(node)
                    curr_qsv_54 = node.qsv
                    return
                elif node.restriction =="55":
                    schedule55.insertAtEnd(node)
                    curr_qsv_55 = node.qsv
                    return
                else:
                    if node.restriction == 'na' and ERT_54 < ERT_55:
                        if cur1 is not None:
                            if cur1.qsv is not node.qsv:
                                ERT_54 += (40 +18*node.noRolls)
                            else:
                                ERT_54 += (18*node.noRolls)
                        schedule54.insertAtEnd(node)
                        curr_qsv_54 = node.qsv
                        return
                    else:
                        if cur2 is not None:
                            if cur2.qsv is not node.qsv:
                                ERT_55 += (40 +18*node.noRolls)
                            else:
                                ERT_55 += (18*node.noRolls)
                        schedule55.insertAtEnd(node)
                        curr_qsv_55 = node.qsv
                        return
 

    if (cur1.qsv != node.qsv) and (cur2.qsv != node.qsv):
        # node.start_time += 40
        if schedule54.findSize(schedule54.head) < schedule55.findSize(schedule55.head):
            if node.restriction == "54":
                schedule54.insertAtEnd(node)
                curr_qsv_54 = node.qsv
                return
            elif node.restriction == '55': 
                schedule55.insertAtEnd(node)
                curr_qsv_55 = node.qsv
                return
            else:
                if node.restriction == 'na' and ERT_54 < ERT_55:
                    if cur1 is not None:
                        if cur1.qsv is not node.qsv:
                                ERT_54 += (40 +18*node.noRolls)
                        else:
                            ERT_54 += (18*node.noRolls)
                    schedule54.insertAtEnd(node)
                    curr_qsv_54 = node.qsv
                    return
                else:
                    if cur1 is not None:
                        if cur2.qsv is not node.qsv:
                                ERT_55 += (40 +18*node.noRolls)
                        else:
                                ERT_55 += (18*node.noRolls)
                    schedule55.insertAtEnd(node)
                    curr_qsv_55 = node.qsv
                    return

        else:
            if node.restriction == "55":
                schedule55.insertAtEnd(node)
                curr_qsv_55 = node.qsv
                return
            elif node.restriction == '54': 
                schedule54.insertAtEnd(node)
                curr_qsv_54 = node.qsv
                return
            else: 
                if node.restriction == 'na' and ERT_54 < ERT_55:
                    if cur1 is not None:
                        if cur1.qsv is not node.qsv:
                                ERT_54 += (40 +10*node.noRolls)
                        else:
                                ERT_54 += (18*node.noRolls)
                    schedule54.insertAtEnd(node)
                    curr_qsv_54 = node.qsv
                    return
                else:
                    if cur1 is not None:
                        if cur2.qsv is not node.qsv:
                                ERT_55 += (40 +18*node.noRolls)
                        else:
                                ERT_55 += (18*node.noRolls)
                    schedule55.insertAtEnd(node)
                    curr_qsv_55 = node.qsv
                    return


    if cur1.qsv == node.qsv:
        findBestSpot(node)
    else:
        findBestSpot(node)

def findBestSpot(node):
    global curr_qsv_54
    global curr_qsv_55
    global ERT_55
    global ERT_54
    cur1 = schedule54.head
    cur2 = schedule55.head
    
    best1 = None
    best2 = None

    while (cur1.next is not None):
        cur1 = cur1.next
    while (cur2.next is not None):
        cur2 = cur2.next
    
    best1 = cur1
    best2 = cur2
    bestImprovement = 0

    while (cur1 is not schedule54.head and cur1.qsv == node.qsv ):
        while (cur2 is not schedule55.head and cur2.end_time > cur1.start_time ):
            if cur2.start_time < cur1.end_time:
                if bestImprovement < isCompatible(cur2, node) - isCompatible(cur1, cur2):
                    bestImprovement = isCompatible(cur2, node) - isCompatible(cur1, cur2)
                    best1 = cur1
                    best2 = cur2            
            cur2 = cur2.prev
        cur1 = cur1.prev
    if bestImprovement <= 0:
        if node.restriction == "54":
            schedule54.insertAtEnd(node)
            curr_qsv_54 = node.qsv
            return
        elif node.restriction == "55": 
            schedule55.insertAtEnd(node)
            curr_qsv_55 = node.qsv
            return
        else: 
            if node.restriction == 'na' and ERT_54 < ERT_55:
                if cur1 is not None:
                    if cur1.qsv is not node.qsv:
                            ERT_54 += (40 +10*node.noRolls)
                    else:
                            ERT_54 += (18*node.noRolls)
                old = schedule54.swap(best1, node)
                schedule54.insertAtEnd(old)
                curr_qsv_54 = node.qsv
                return
            else:
                if cur1 is not None:
                    if cur2.qsv is not node.qsv:
                            ERT_55 += (40 +18*node.noRolls)
                    else:
                            ERT_55 += (18*node.noRolls)
                old = schedule55.swap(best2, node)
                schedule55.insertAtEnd(old)
                curr_qsv_55 = node.qsv
                return
    else:
        if node.restriction == "54":
            old = schedule54.swap(best1, node) 
            schedule54.insertAtEnd(old)
            return
        elif node.restriction == "55": 
            old = schedule55.swap(best2, node) 
            schedule55.insertAtEnd(old)
            return
        else: 
            if node.restriction == 'na' and ERT_54 < ERT_55:
                if cur1 is not None:
                    if cur1.qsv is not node.qsv:
                            ERT_54 += (40 +10*node.noRolls)
                    else:
                            ERT_54 += (18*node.noRolls)
                old = schedule54.swap(best1, node)
                schedule54.insertAtEnd(old)
                curr_qsv_54 = node.qsv
                return
            else:
                if cur1 is not None:
                    if cur2.qsv is not node.qsv:
                            ERT_55 += (40 +18*node.noRolls)
                    else:
                            ERT_55 += (18*node.noRolls)
                old = schedule55.swap(best2, node)
                schedule55.insertAtEnd(old)
                curr_qsv_55 = node.qsv
                return
                
        # if node.restriction == "55":
        #     schedule55.insertAtEnd(node)
        #     curr_qsv_55 = node.qsv
        #     return
        # else: 
        #     schedule54.insertAtEnd(node)
        #     curr_qsv_54 = node.qsv
        #     return

def isCompatible(node1, node2):
    compatibiltyScore = 0
    totalNoLanes = node1.noLanes + node2.noLanes
    if totalNoLanes > 16:
        compatibiltyScore -= 2
    elif totalNoLanes > 14 and totalNoLanes <= 16:
        compatibiltyScore -= 1
    elif totalNoLanes >= 12 and totalNoLanes <= 14:
        compatibiltyScore += 1
    else:
        compatibiltyScore += 2

    if node1.isCoprint is True and node2.isCoprint is True:
        if node1.no_coprint + node2.no_coprint >= 8:
            compatibiltyScore -= 2
        else:
            compatibiltyScore -= 1
    if ((node1.isCoprint is True and node2.isLarge is True) or (node2.isCoprint is True and node1.isLarge is True)):
        compatibiltyScore += 1

    totalWaste = node1.waste + node2.waste
    if totalWaste > 1500:
        compatibiltyScore -= 2
    elif totalWaste >= 500 and totalWaste <= 1500:
        compatibiltyScore -= 1
    elif totalWaste < 500:
        compatibiltyScore += 1

    return compatibiltyScore
                
for index, row in grouped_df.iterrows():
   new_node = create_node(row)
   scanToInsert(new_node)

# for index, row in sorted_b2.iterrows():
#    new_node = create_node(row)
#    scanToInsert(new_node)

# for index, row in sorted_b3.iterrows():
#    new_node = create_node(row)
#    scanToInsert(new_node)

# for index, row in sorted_wr_df.iterrows():
#    new_node = create_node(row)
#    scanToInsert(new_node)

sorted_wr_df
# print('\nSchedule 54')
print("orderID,qsv,noLanes,noRolls,waste,no_coprint,restriction,due_date", file=file1)
print(schedule54.printList(schedule54.head), file=file1)
# print('\nSchedule 55')
print("orderID,qsv,noLanes,noRolls,waste,no_coprint,restriction,due_date", file=file2)
print(schedule55.printList(schedule55.head), file=file2)

file1.close()

df_54 = pd.read_csv("output1.txt", delimiter=",")
df_54["package size"] = None
df_54["approximate duration"] = None
df_54["customer"] = None
for index, rows in df_54.iterrows():
    matching_id = sp_df[sp_df['Order number']==rows['orderID']]
    volume = matching_id.iloc[0,2]
    shape = matching_id.iloc[0,3]
    customer = matching_id.iloc[0,6]
    df_54.at[index, 'package size'] = str(volume)+"-"+str(shape)
    df_54.at[index, 'approximate duration'] = rows["noRolls"]*18
    df_54.at[index, 'customer'] = customer

# column_order = ["orderID", "qsv", "package size", "noLanes", "noRolls", "approximate duration", "waste", "no_coprint", "restriction", "due_date"]

df_54 = df_54[["orderID", "customer","qsv", "package size", "noLanes", "noRolls", "approximate duration", "waste", "no_coprint", "restriction", "due_date"]]

total_dict = {"orderID":["","Estimated Total Run Time: ",""], 
            "package size": ["", ERT_54, ""]}
temp_df = pd.DataFrame(total_dict)
# print(temp_df)
df_54 = pd.concat([df_54, temp_df], ignore_index = True)

# print(df_54)
       
file2.close()

df_55 = pd.read_csv("output2.txt", delimiter=",")
df_55["package size"] = None
df_55["approximate duration"] = None
df_55["customer"] = None
for index, rows in df_55.iterrows():
    matching_id = sp_df[sp_df['Order number']==rows['orderID']]
    volume = matching_id.iloc[0,2]
    shape = matching_id.iloc[0,3]
    customer = matching_id.iloc[0,6]
    df_55.at[index, 'package size'] = str(volume)+"-"+str(shape)
    df_55.at[index, 'approximate duration'] = rows["noRolls"]*18
    df_55.at[index, 'customer'] = customer

df_55 = df_55[["orderID", "customer", "qsv", "package size", "noLanes", "noRolls", "approximate duration", "waste", "no_coprint", "restriction", "due_date"]]
total_dict = {"orderID":["","Estimated Total Run Time: ",""], 
            "package size": ["", ERT_55,""]}
temp_df = pd.DataFrame(total_dict)
# print(temp_df)
df_55 = pd.concat([df_55, temp_df], ignore_index = True)


# print(df_55)

root = Tk()
root.withdraw()
folder_selected = filedialog.askdirectory(title = "Select Output Folder")
print(folder_selected)
output_filename = folder_selected+"/schedules.xlsx"

writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
df_54.to_excel(writer, sheet_name='Schedule 54', index = False)
df_55.to_excel(writer, sheet_name='Schedule 55', index = False)


tmp_54 = df_54
tmp_54["Scheduled Machine"] = 54
tmp_54.drop(tmp_54.tail(3).index,inplace=True)
total_dict = {"orderID":["","Estimated Total Run Time: ",""], 
            "package size": ["", ERT_54,""]}
temp_df = pd.DataFrame(total_dict)
# print(temp_df)
tmp_54 = pd.concat([tmp_54, temp_df], ignore_index = True)

tmp_55 = df_55
tmp_55["Scheduled Machine"] = 55
tmp_55.drop(tmp_55.tail(3).index,inplace=True)
total_dict = {"orderID":["","Estimated Total Run Time: ",""], 
            "package size": ["", ERT_55,""]}
temp_df = pd.DataFrame(total_dict)
# print(temp_df)
tmp_55 = pd.concat([tmp_55, temp_df], ignore_index = True)

blank= {"":[""]}
tmp_blank = pd.DataFrame(blank)
frames = [tmp_54, tmp_blank, tmp_55]
df_combined = pd.concat(frames, axis=1)

df_combined.to_excel(writer, sheet_name='Combined Schedules', index = False)

# os.remove('output.txt')
os.remove('output1.txt')
os.remove('output2.txt')

print("Generated slitter schedules in file 'schedules.xlsx'")
print( "Estimated Run Time on Slitter 54:", ERT_54, "minutes")
print( "Estimated Run Time on Slitter 55:", ERT_55, "minutes")

writer.save()
# %%
