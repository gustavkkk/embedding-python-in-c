import lxml
import docx

def HelloWorld():
    print("Hello World")

def add(a, b):
    return a+b
	
def TestDict(dict):
    print(dict)
    dict["Age"] = 17
    return dict

class Person:  
    def greet(self, greetStr):
        print(greetStr)

def multiply(a,b):
    print("Will compute", a, "times", b)
    c = 0
    for i in range(0, a):
        c = c + b
    return c