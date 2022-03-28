
import pandas as pd
import numpy as np
import tkinter as tk 

from tkinter import filedialog


# Defining data structures for doubly linked lists 

class Node: 
    def __init__(self, orderID, isCoprint, noRolls, noLanes): 
        self.orderID = orderID
        self.isCoprint = isCoprint
        self.noRolls = noRolls
        self.noLanes = noLanes
        self.next = None 
        self.prev = None 

class doublyLinkedList: 
    def __init__(self): 
        self.head = None

    def insertAtBeg(self, orderID, isCoprint, noRolls, noLanes): 

        new_node = Node(orderID, isCoprint, noRolls, noLanes)
        new_node.next = self.head
        new_node.prev = None 


        if self.head is not None: 
            self.head.prev = new_node

        self.head = new_node


    def insertAtPos(self, prev_node, orderID, isCoprint, noRolls, noLanes):
        if prev_node is None: 
            print("This node does not exist")
            return 

        new_node = Node(orderID, isCoprint, noRolls, noLanes) 
        new_node.next = prev_node.next
        prev_node.next = new_node
        new_node.prev = prev_node

        if new_node.next is not None: 
            new_node.next.prev = new_node
            

    def insertAtEnd(self, orderID, isCoprint, noRolls, noLanes): 
        new_node = Node(orderID, isCoprint, noRolls, noLanes)
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
 
        print("\nTraversal in forward direction")
        while node:
            print(" {}".format(node.orderID))
            last = node
            node = node.next
    


# llist = doublyLinkedList()
# llist.insertAtEnd("36", 4, 5, 3)
# llist.insertAtEnd("37", 4, 5, 3)
# llist.insertAtEnd("39", 4, 5, 3)
# llist.insertAtEnd("42", 4, 5, 3)
# llist.insertAtBeg("64", 4, 5, 3)
# llist.insertAtBeg("93", 4, 5, 3)
# llist.insertAtBeg("72", 4, 5, 3)

# llist.printList(llist.head)

tk.Tk().withdraw() # prevents an empty tkinter window from appearing

#asking the user to select their input file
file_path = filedialog.askopenfilename()

print(file_path)

#converting input excel file into a dataframe 
df = pd.read_csv(file_path)

print(df)

#cleaning up input file 
df = df.dropna(axis=0)
df.columns = df.iloc[0]
df = df.drop(df.index[0])

#print(df)

# we only care about WIP's that are in the warehouse after the lamination stage ('WIP Printing-Lamina')
df = df[df['Material']=='WIP Printing-Lamina']
#print(df)

# filtering the dataframe to only show values which we want for current rules implementation (Order, Lanes, No of Rolls, POrder Due Date)
df = df[[" Order", " Lanes", " No Of Rolls", " POrder Due Date"]]
df = df.astype({" Order": str, " Lanes": int, " No Of Rolls": int})
df[" POrder Due Date"] = pd.to_datetime(df[" POrder Due Date"])
print(df)
#print(df.dtypes)

# sorting the data set by due date 

sorted_df = df.sort_values(by = [" POrder Due Date", " No Of Rolls"])
print(df[:20])

schedule54 = doublyLinkedList() # linked list maintaining schedule for slitter 54
schedule55 = doublyLinkedList() # linkedlist maintaining schedule for slitter 55

# reading through the dataframe line by lone to create nodes 
for index, row in sorted_df.iterrows():
    orderID = row[' Order']
    noLanes = row[' Lanes']
    noRolls = row[' No Of Rolls']
    dueDate = row[' POrder Due Date']
    current_node = [orderID, noLanes, noRolls, dueDate]
    #print(current_node)



# Implementation of current 5 scheduling rules




