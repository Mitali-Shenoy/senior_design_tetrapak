
from re import T
import pandas as pd
import numpy as np
import tkinter as tk 

from tkinter import filedialog

file = open('output.txt', 'w')

# Defining data structures for doubly linked lists 

class Node: 
    def __init__(self, orderID, noRolls, noLanes): 
        self.orderID = orderID
#        self.isCoprint = isCoprint
        if noRolls > 3:
            self.isLarge = True
        else:
            self.isLarge = False
        if orderID[-1] == 'C':
            self.isCoprint = True
        else:
            self.isCoprint = False
        self.noRolls = noRolls
        self.noLanes = noLanes
        self.next = None 
        self.prev = None 

class doublyLinkedList: 
    def __init__(self): 
        self.head = None

    def insertAtBeg(self, orderID, noRolls, noLanes): 

        new_node = Node(orderID, noRolls, noLanes)
        new_node.next = self.head
        new_node.prev = None 


        if self.head is not None: 
            self.head.prev = new_node

        self.head = new_node


    def insertAtPos(self, prev_node, orderID, noRolls, noLanes):
        if prev_node is None: 
            print("This node does not exist")
            return 

        new_node = Node(orderID, noRolls, noLanes) 
        new_node.next = prev_node.next
        prev_node.next = new_node
        new_node.prev = prev_node

        if new_node.next is not None: 
            new_node.next.prev = new_node
            

    def insertAtEnd(self, orderID, noRolls, noLanes): 
        new_node = Node(orderID, noRolls, noLanes)
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
            print(" {}".format(node.orderID), file=file)
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



def comparison(current):
    if schedule54.head is None:
        schedule54.insertAtEnd(current.orderID, current.noRolls, current.noLanes)
        return
    elif schedule55.head is None:
        schedule55.insertAtEnd(current.orderID, current.noRolls, current.noLanes)
        return
    else:
        last54 = schedule54.head
        last55 = schedule55.head
        while (last54.next is not None): 
            last54 = last54.next
        while (last55.next is not None): 
            last55 = last55.next
        if compatible(last54, last55) is False:
            if compatible(last54, current) is True:
                schedule54.insertAtPos(last55.prev, current.orderID, current.noRolls, current.noLanes)
                return
            if compatible(last55, current) is True:
                schedule55.insertAtPos(last54.prev, current.orderID, current.noRolls, current.noLanes)
                return
        if schedule54.findSize(schedule54.head) > schedule55.findSize(schedule55.head):
            schedule55.insertAtEnd(current.orderID, current.noRolls, current.noLanes)
        else:
            schedule54.insertAtEnd(current.orderID, current.noRolls, current.noLanes)
        return

# Implementation of current 5 scheduling rules

def compatible(a, b):
    if (a.isLarge == True and b.isCoprint == True) or (b.isLarge == True and a.isCoprint == True):
        return True
    elif a.noLanes == 9 and b.noLanes == 9:
        return False
    elif a.isCoprint == True and b.isCoprint == True:
        return False
    else:
        return True

# reading through the dataframe line by lone to create nodes 
for index, row in sorted_df.iterrows():
    orderID = row[' Order']
    noLanes = row[' Lanes']
    noRolls = row[' No Of Rolls']
#    dueDate = row[' POrder Due Date']
    current_node = Node(orderID, noLanes, noRolls)
    #print(current_node)
    for i in range(noRolls):
        comparison(current_node)      


print("Schedule for slitter 54:", file=file)        
schedule54.printList(schedule54.head)
print("\nSchedule for slitter 55:", file=file)
schedule55.printList(schedule55.head)
file.close()


