from oopAnotherModule import anotherVariable # we are tapping OAM and getting from inside anotherVariable from the class
print(anotherVariable) # since we took it from the other place we can print what we called for
from xlwingsBasic import channelNumsSliced
from turtle import Turtle, Screen # we are importing all of the Turtle class now to summon whats in the class with a variable declared

speedy = Turtle()
myScreen = Screen()

speedy.shape("turtle")
speedy.shapesize(3,3,4)
speedy.color('red')
speedy.pencolor('blue')
speedy.pensize(22)

speedy.pendown()
speedy.forward(300)
print(channelNumsSliced)

myScreen.exitonclick()

