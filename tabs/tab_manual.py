import dash_core_components as dcc

MD = '''
# INSTRUCTION

### Optimization Model
Upload
- Download input template
- Input data in file
- Upload
- Check upload details

Validate
- Sheet validation
  - Columns - check complete columns in upload file for all sheets
  - Duplicate - check there is duplicate value of master columns for all sheets
- Feasible validation
  - Validate1 - check if all demand have supply chain parameters
  - Validate2 - check supplyprod cap must be more than demand (by product)
  - Validate3 - check logistics cap must be more than demand (by destination)
  - Validate4 - check logistics cap vs supply cap (by supply)
  - Validate5 - check logistics cap vs supply cap (by supply)

Solve
- Press solve button to solve the problem
- Check solving status
  - Status refer to solver status ('ok' = complete)
  - Condition refer to termination condition ('optimal' = solution is optimal)
  - Please find reference [here](http://www.pyomo.org/blog/2015/1/8/accessing-solver)
- Download output file

### Visualization
- Press refresh to refresh the latest output
- Check details of the file
- Click on supply chart (bar chart) to filter all pages by selected supply
- Click other area in supply chart to reset

# REFERENCE

### Optimization model

Mixed Integer Linear Programming (MIP)

Dimensions: supply, product, route, warehouse, destination

Objective: Find the optimized supply chain production and route to serve demand of customers by minimize net contribution of supply chain parameters

Constraints:
- Supply capacity (min/max)
- Supply product cap
- Warehouse fixed cost

### Tools
- Code: [Python](https://www.python.org/)
- Optimization Libraries: [Pyomo](http://www.pyomo.org/)
- Solver Engine: [CBC](https://projects.coin-or.org/Cbc) / [GLPK](https://www.gnu.org/software/glpk/)
- Front-end: [Dash](https://dash.plot.ly/)
'''

tab = [
    dcc.Markdown(MD)
]
