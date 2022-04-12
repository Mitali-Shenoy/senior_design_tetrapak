
# %%
from re import T
import pandas as pd
import numpy as np
import tkinter as tk 

from tkinter import filedialog

# %%

file = open('output.txt', 'w')
# Defining data structures for doubly linked lists 

class Node: 
    def __init__(self, orderID, noRolls, noLanes, qsv, waste, no_coprint, restriction): 
        
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
            

    def insertAtEnd(self, new_node): 
        # new_node = Node(orderID, noRolls, noLanes)
        last = self.head

        new_node.next = None

        if self.head is None: 
            new_node.prev = None 
            self.head = new_node
            return 

        while (last.next is not None): 
            last = last.next
        
        last.next = new_node 
        new_node.prev = last

    #def get():
    
    def printList(self, node):
 
#        print("\nTraversal in forward direction")
        while node:
            
            print(" {}, {}, {}, {}, {}, {}, {}".format(node.orderID, node.noRolls, node.noLanes, node.qsv, node.waste, node.no_coprint, node.restriction), file=file)
            last = node
            node = node.next
    
    # This function returns size of linked list
    def findSize(self, node):
    
        res = 0
        while (node != None):
            res = res + 1
            node = node.next
        
        return res
    


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
file = open('output.txt', 'w')

tk.Tk().withdraw() # prevents an empty tkinter window from appearing

#asking the user to select their input file
file_path = filedialog.askopenfilename()

print(file_path)

#converting input excel file into a dataframe 
wr_df = pd.read_excel(file_path) #read Wip Report File (wp)

print(wr_df)


file = open('output.txt', 'w')

tk.Tk().withdraw() # prevents an empty tkinter window from appearing

#asking the user to select their input file
file_path = filedialog.askopenfilename()

print(file_path)

# converting input excel file into a dataframe 


sp_df = pd.read_excel(file_path) #read Slit Plan file

print(sp_df)


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
print(wr_df)

print()
# print(sp_df.columns)

# cleaning up the in[ut data file for Slit Plan
sp_df = sp_df[[0, "Order number", "Package.Volume", "Package.Shape", "Lane Assignment"]]
sp_df = sp_df.astype({0 : int, "Order number": str, "Package.Volume": str, "Package.Shape": str, "Lane Assignment": str}) #setting the datatype for columns

print(sp_df)

# %%

sorted_wr_df = wr_df.sort_values(by = [" POrder Due Date"])
# print(sorted_wr_df[:20])

order_len = len(sorted_wr_df)
bucket_1 = sorted_wr_df.iloc[0:int(order_len/3),:] 
bucket_2 = sorted_wr_df.iloc[int((order_len/3)):int(2*(order_len/3))+1,:] 
bucket_3 = sorted_wr_df.iloc[int((2*(order_len/3))+1):order_len+1,:] 


print("\nBucket 1")
sorted_b1 = bucket_1.sort_values(by = [" QSV"])
print(sorted_b1)
print("Bucket 2\n")
sorted_b2 = bucket_2.sort_values(by = [" QSV"])
print(sorted_b2)
print("Bucket 3\n")
sorted_b3 = bucket_3.sort_values(by = [" QSV"])
print(sorted_b3)


# %%

schedule54 = doublyLinkedList() # linked list maintaining schedule for slitter 54
schedule55 = doublyLinkedList() # linkedlist maintaining schedule for slitter 55

#creating nodes from the input file 

print()
print()

def create_node(wip_row):

    orderID = wip_row[" Order"]
    noRolls = wip_row[" No Of Rolls"]
    noLanes = wip_row[" Lanes"]
    qsv = wip_row[" QSV"]
    waste = wip_row[" Potential Waste Length"]

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


    # print(orderID, noRolls, noLanes, qsv, waste, no_coprint, restriction)
    new_node = Node(orderID, noRolls, noLanes, qsv, waste, no_coprint, restriction)

# for index, row in sorted_b1.iterrows():
#     create_node(row)