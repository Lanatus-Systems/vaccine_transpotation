import pandas as pd
import pulp
from dash import Dash, html, dash_table

# Data Preparation
rate_excel = pd.read_excel("Rate.xlsx")
rate_data = rate_excel.rename(columns=lambda x: x.strip())

variable_excel = pd.read_excel("Var.xlsx")
city_mode_data = variable_excel.rename(columns=lambda x: x.strip())

transportation_excel = pd.read_excel("Mode3.xlsx")
transportation_excel = transportation_excel.rename(columns=lambda x: x.strip())
mode_data = transportation_excel.rename(columns={
    'Mode': 'Mode',
    'Capacity (number of Doses can carry)': 'Capacity',
    'Speed(KM/Hr)': 'Speed',
    'Cost per KM($) (fuel, driver, maintenance)': 'Cost per KM',
    'Number of Vehicles (Daily)': 'Number of Vehicles',
    'Cost per dose per KM': 'Cost per dose per KM',
    'Max Capaity': 'Max Capacity'
})

infection_excel = pd.read_excel("Infection.xlsx")
infection_excel = infection_excel.rename(columns=lambda x: x.strip())
city_demand = infection_excel.rename(columns={
    'City': 'City',
    'Infection rate': 'Infection rate',
    'Doses Needed (Daily)': 'Doses Needed'
})
city_demand['City'] = city_demand['City'].str.split(',').str[0].str.strip()

city_excel = pd.read_excel("City.xlsx")
city_excel = city_excel.rename(columns=lambda x: x.strip())
distance_data = city_excel.rename(columns={
    'Source - Plant': 'Source',
    'Destination': 'Destination',
    'Distance: Source to destination(KM)': 'Distance',
    'Time_Truck  (Distnace/Speed)': 'Cost Truck',
    'Time_Train (Distnace/Speed)': 'Cost Train',
    'Time_Plane (Distance/Speed)': 'Cost Aeroplane'
})
distance_data['Destination'] = distance_data['Destination'].str.split(',').str[0].str.strip()

# Define tables for display
tables = [
    ("Rate Data", "Sheet1", rate_data),
    ("City Mode Data", "Sheet2", city_mode_data),
    ("Mode Data", "Sheet3", mode_data),
    ("City Demand", "Sheet4", city_demand),
    ("Distance Data", "Sheet5", distance_data)
]

max_tranport_cap = 2000000 

# Optimization Model
def run_optimization():
    prob = pulp.LpProblem("Vaccine_Delivery", pulp.LpMaximize)

    #dicition var for Transport_(city, mode)
    transport_vars = pulp.LpVariable.dicts("Transport",
                                         ((city, mode) for city in city_demand['City'] 
                                          for mode in mode_data['Mode']),
                                         lowBound=0, cat='Integer')
    print("transport_vars ====>",transport_vars)
    
    #dictionary for priority score city wise 
    priority_scores = dict(zip(city_demand['City'], city_demand['Infection rate']))
    print("priority_scores ======>",priority_scores)

    weight_mapping = {1: 0.8, 2: 0.4, 3: 0.1}
    #objective function
    prob += pulp.lpSum([transport_vars[(city, mode)] * mode_data[mode_data['Mode'] == mode]['Capacity'].values[0] *
                        weight_mapping[city_demand[city_demand['City'] == city]['Infection rate'].values[0]]
                       for city in city_demand['City'] 
                       for mode in mode_data['Mode']])
    print("third =====>",pulp.lpSum([transport_vars[(city, mode)] * mode_data[mode_data['Mode'] == mode]['Capacity'].values[0] *
                        weight_mapping[city_demand[city_demand['City'] == city]['Infection rate'].values[0]]
                       for city in city_demand['City'] 
                       for mode in mode_data['Mode']]))
    
    #vehicle constraint
    for mode in mode_data['Mode']:
        prob += pulp.lpSum([transport_vars[(city, mode)] 
                            for city in city_demand['City']]) <= mode_data[mode_data['Mode'] == mode]['Number of Vehicles'].values[0]
        print("forth===>",pulp.lpSum([transport_vars[(city, mode)] 
                            for city in city_demand['City']]) <= mode_data[mode_data['Mode'] == mode]['Number of Vehicles'].values[0])

    #max tranportation limit
    prob += pulp.lpSum([transport_vars[(city, mode)] * mode_data[mode_data['Mode'] == mode]['Capacity'].values[0]
                            for city in city_demand['City']
                            for mode in mode_data['Mode']]) <= max_tranport_cap
    print("fifth =====>",pulp.lpSum([transport_vars[(city, mode)] * mode_data[mode_data['Mode'] == mode]['Capacity'].values[0]
                            for city in city_demand['City']
                            for mode in mode_data['Mode']]) <= max_tranport_cap)
    
    #city demand constraint
    for city in city_demand['City']:
        prob += pulp.lpSum([transport_vars[(city, mode)] *
                            mode_data[mode_data['Mode'] == mode]['Capacity'].values[0]
                                for mode in mode_data['Mode']]) <= city_demand[city_demand['City'] == city]['Doses Needed'].values[0]
        print(pulp.lpSum([transport_vars[(city, mode)] *
                            mode_data[mode_data['Mode'] == mode]['Capacity'].values[0]
                                for mode in mode_data['Mode']]) <= city_demand[city_demand['City'] == city]['Doses Needed'].values[0])
        
    prob.solve()

    results = []
    for city in city_demand['City']:
        for mode in mode_data['Mode']:
            if pulp.value(transport_vars[(city, mode)]) and pulp.value(transport_vars[(city, mode)]) > 0:
                doses = pulp.value(transport_vars[(city, mode)]) * mode_data[mode_data['Mode'] == mode]['Capacity'].values[0]
                cost = doses * distance_data[distance_data['Destination'] == city][f'Cost {mode}'].values[0]
                results.append({
                    'City': city,
                    'Mode': mode,
                    'Vehicles': pulp.value(transport_vars[(city, mode)]),
                    'Doses Delivered': doses,
                    'Total Cost': cost
                })
    return pd.DataFrame(results)

# Run optimization
results_df = run_optimization()

# Dash Application
app = Dash(__name__)

app.layout = html.Div([
    html.H1("Vaccine Delivery Optimization"),
    
    html.H2("Input Data"),
    html.Div([
        html.Div([
            html.H3(f"{file}"),
            dash_table.DataTable(
                data=df.to_dict('records'),
                columns=[{"name": i, "id": i} for i in df.columns],
                page_size=10,
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'left'},
                style_header={'backgroundColor': '#f0f0f0', 'fontWeight': 'bold'},
            ),
            html.Hr()
        ]) for file, sheet, df in tables
    ]),
    
    html.H2("Optimization Results"),
    dash_table.DataTable(
        data=results_df.to_dict('records'),
        columns=[{'name': i, 'id': i} for i in results_df.columns],
        style_table={'overflowX': 'auto'},
        style_data_conditional=[{
            'if': {'row_index': 'odd'},
            'backgroundColor': 'rgb(248, 248, 248)'
        }]
    )
])

if __name__ == '__main__':
    app.run(debug=True)