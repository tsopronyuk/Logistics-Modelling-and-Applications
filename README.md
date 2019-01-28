# Logistics-Modelling-and-Applications

## VBA Project

### Logistics Modelling and Applications

Project Specifications:
Vehicle routing problem (VRP) is an important area in logistics and has many practical applications. The objective of VRP is to find the optimal set of routes for a fleet of vehicles to traverse in order to deliver to a given set of customers. Travelling Salesman Problem (TSP) is a variant of VRP: given a distance matrix among a set of customers a Salesman starts from one customer, visits each customer once and exactly once and returns to the starting customer. The objective of TSP is to find the sequence of visits that minimises the total distance the Salesman travelled.

The data set is given in att48_data.xlsx file.

The model should be able to perform four primary tasks:

1.	Get model input: Get distance data ready for your model. You can either copy the given data to a worksheet in your model, then read the cell values from VBA code; or you may use VBA to read the given data from its source file.

2.	Construct an initial route using random tour algorithm in VBA. 

a.	Randomly generate 3 tours, say n.

b.	Find the shortest possible route from routes generated in task 2a.

3.	Output the chosen sequence and the total distance of the best initial solution.

4.	Change the value of n in task 2a (e.g. n=5,10, 20, 50,100 random tours) and run the VBA macro again. Will the shortest total distance and sequence of your best tour be different? Please analyse the result in your technical report and display both the best and the worst tour in excel. 

You can embellish your application by using buttons, charts and forms according to your personal taste. You can also add other functionalities that you consider useful (e.g., you may use userform to get the value of n as an input from user, you may allow the user to use a default setting by clicking a button.) Initiative and creativity will be rewarded.



