from random import *

class Node :
    def __init__(self, x, y, m):
        #provjera jeli provjeren
        #self.checked = 0
        self.x = x
        self.y = y
        self.matrix = []
        for i in range(m):
            col = []
            for j in range(m):
                col.append(randint(0,1))
            self.matrix.append(col)

class Network:
    def __init__(self, n, m):
        #spremljene dimenzije mreze
        self.n = n
        self.network = []
        for i in range(n):
            col = []
            for j in range(n):
                node = Node(i, j, m)
                col.append(node)
            self.network.append(col)

    def neighbours(self, x, y):
        neighbour = []

        #print print_node(self.network[int(x[0])][int(y[0])])
        
        if((x - 1) > -1) :
            n1 = self.network[(x - 1)][y]
            #print "Gornji susjed : "
            #print_node(n1)
            neighbour.append(n1)
        
        if((y + 1) < self.n) :
            n2 = self.network[x][(y + 1)]
            #print "Desni susjed : "
            #print_node(n2)
            neighbour.append(n2)
        
        if((x + 1) < self.n) :
            n3 = self.network[(x + 1)][y]
            #print "Donji susjed : "
            #print_node(n3)
            neighbour.append(n3)
        
        if((y - 1) > - 1) :
            n4 = self.network[x][(y - 1)]
            #print "Lijevi susjed : "
            #print_node(n4)
            neighbour.append(n4)

        return neighbour
            

def print_node(temp):
    for i in range(len(temp.matrix)):
            print temp.matrix[i]

    
def print_network(temp):
    for i in range(len(temp.network[0])):
        for k in range(len(temp.network[i][0].matrix[0])):
            for j in range(len(temp.network[i])):             
                print str(temp.network[i][j].matrix[k])+"\t",
            print '\n'
        print '\n'
    
def check_knowledge(node):
    for i in range(len(node.matrix[0])):
        for j in range(len(node.matrix[i])):
            if(node.matrix[i][j] == 0) : return False
    return True


def replace_knowledge(node, neighbour):
    for i in range(len(node.matrix[0])):
        for j in range(len(node.matrix[i])):
            if(node.matrix[i][j] == 0):
                for temp in neighbour:
                    if(temp.matrix[i][j] == 1) :
                        node.matrix[i][j] = temp.matrix[i][j]
                        break
    return node


def recursion_knowledge(node, network):
    if(check_knowledge(node)) : return
    #treba dodati provjeru ako nema jedinica da napune da izbaci. znaci nekako naci kraj tih rekurzija.
    neighbour = network.neighbours(node.x,node.y)
    replace_knowledge(node, neighbour)

    for neighbour_node in neighbour:
        recursion_knowledge(neighbour_node, network)
                 
    
def main():
    #network n x n
    n = 2
    #node m x m
    m = 6
    #stvaranje, inicijaliziranje i printanje mreze
    _network = Network(n,m)
    print print_network(_network)

    #odabir i ispis cvora
    while(True):
        x = input("Odaberite redak koordinatu cvora : ")
        if (x > -1 and x < n) : break
    while(True):
        y = input("Odaberite stupac koordinatu cvora : ")
        if (y > -1 and y < n) : break
    selected_node = _network.network[x][y]
    print "Odabran cvor : "
    print print_node(selected_node)
        
    recursion_knowledge(selected_node, _network)

    print "Nakon promjene : "
    print_node(selected_node)

if __name__=='__main__':
    main()
            
