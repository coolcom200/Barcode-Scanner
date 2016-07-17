from matplotlib import pyplot
from random import randint

List = []
for i in range(30):
    x = randint(30, 40)
    List.append(x)
pyplot.plot(List)
pyplot.ylabel('Attendance')
pyplot.show()
