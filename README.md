# Installation

pip install -r requirements.txt

install optimization engine
- [CBC](https://projects.coin-or.org/Cbc)
- [GLPK](https://www.gnu.org/software/glpk/)

create config.yaml file

# Optimization model

Mixed Integer Linear Programming (MIP)

Dimensions: supply, product, route, warehouse, destination

Objective: Find the optimized supply chain production and route to serve demand of customers by minimize net contribution of supply chain parameters

Constraints:
- Supply capacity (min/max)
- Supply product cap
- Warehouse fixed cost

# Tools
- Code: [Python](https://www.python.org/)
- Optimization Libraries: [Pyomo](http://www.pyomo.org/)
- Solver Engine: [CBC](https://projects.coin-or.org/Cbc) / [GLPK](https://www.gnu.org/software/glpk/)
- Front-end: [Dash](https://dash.plot.ly/)
