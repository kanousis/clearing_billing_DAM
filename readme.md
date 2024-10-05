<h1> Clearing and Billing of Day-Ahead Market using linear programming </h1>

In this project we solve the clearing and pricing problem of the Greek Day-Ahead Market (ignoring the limitations of the transmission network and the integration characteristics of conventional units).

On the website of the Hellenic Energy Exchange (specifically at the link https://www.enexgroup.gr/web/guest/markets-publications-el-day-ahead-market and in the “Aggregated Buy/Sell Orders Curves” section) there are published per day the offers placed by the participants in the Greek daily electricity market (in one excel file per day). 

Solution steps:

We download the file of the day we are interested in solving the problem

We then split the orders in 24 different excel files (input_files), each corresponding to each hour of the following day. (creating_input_output_files.py)

The file code_pupl.py accepts as input one by one the 24 files we created above and solves the optimization problem for each one. 
We used Python's PuLP library, specifically the CBC solver.
Finally we enter our results in the corresponding 24 output files (output_files).

The file plots.py creates the graphs of the results


