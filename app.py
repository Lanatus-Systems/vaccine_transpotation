import pandas as pd
import pulp
from dash import Dash, html, dash_table
import dash_bootstrap_components as dbc

# Data Preparation
rate_excel = pd.read_excel("Rate.xlsx")
rate_data = rate_excel.rename(columns=lambda x: x.strip())
rate_data = rate_data.rename(columns={
    'Weight (>0 and <=1)': 'weight',
    'Value': 'Value',
    'Rate': 'Rate',
    'Doses': 'Doses'
})

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
    'Cost per dose per KM(Cost/Capacity)': 'Cost per dose per KM',
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
    'Time_Plane (Distance/Speed)': 'Cost Airplane'
})
distance_data['Destination'] = distance_data['Destination'].str.split(',').str[0].str.strip()

city_demand2 = infection_excel.rename(columns={
    'City': 'City',
    'Infection rate': 'Infection rate',
    'Doses Needed (Daily)': 'Doses Needed'
})
city_demand2['City'] = city_demand2['City'].str.split(',').str[0].str.strip()

data = {
    'City': distance_data['Destination'],
    'Population': city_demand['Doses Needed'],
    'Infected': [0,0,0,0,0,0],
    'Uninfected': city_demand['Doses Needed'],
    'Total_Vaccinated': [0,0,0,0,0,0]
}

# Create Table
updated_df = pd.DataFrame(data)

# Define tables for display
tables = [
    ("Infection Rate", "Sheet1", rate_data),
    ("City Mode Varaible", "Sheet2", city_mode_data),
    ("Transportation Mode", "Sheet3", mode_data),
    ("Infection Data", "Sheet4", city_demand2),
    ("City", "Sheet5", distance_data),
    # ("updated Data","Sheet6",updated_df)
]
max_production = 120000 
last_day = 0
total = sum(city_demand['Doses Needed'])

# Optimization Model
def run_optimization():
    results = []
    day = 1
    summision = 0
    mincapacity = min(mode_data['Capacity'])
    custom = 0
    max_production = 120000
    if(sum(mode_data['Max Capacity'])<max_production):
        max_production = sum(mode_data['Max Capacity'])

    while(sum(updated_df['Uninfected'])>0):
        for index,row in updated_df.iterrows():
            if(updated_df.at[index,'Uninfected']<mincapacity):
                custom += updated_df.at[index,'Uninfected']
                updated_df.at[index,'Uninfected'] = 0

        prob = pulp.LpProblem("Vaccine_Delivery", pulp.LpMaximize)

        #dicition var for Transport_(city, mode)
        transport_vars = pulp.LpVariable.dicts("Transport",
                                            ((city, mode) for city in city_demand['City'] 
                                            for mode in mode_data['Mode']),
                                            lowBound=0, cat='Integer')

        rate_weight = dict(zip(rate_data['Value'], rate_data['weight']))
        priority_score = {}
        for index,row in city_demand.iterrows():
            city = row['City']
            infection_rate = row['Infection rate']
            doses_need = row['Doses Needed']
            if(doses_need != 0):
                priority_score[city] = (rate_weight[infection_rate]*100000 / doses_need)
            else:
                priority_score[city] = 0

        #objective function
        prob += pulp.lpSum([transport_vars[(city, mode)] * mode_data[mode_data['Mode'] == mode]['Capacity'].values[0] *
                            priority_score[city]
                        for city in city_demand['City'] 
                        for mode in mode_data['Mode']])
    
        #vehicle constraint
        for mode in mode_data['Mode']:
            prob += pulp.lpSum([transport_vars[(city, mode)] 
                                for city in city_demand['City']]) <= mode_data[mode_data['Mode'] == mode]['Number of Vehicles'].values[0]

        #max tranportation limit
        prob += pulp.lpSum([transport_vars[(city, mode)] * mode_data[mode_data['Mode'] == mode]['Capacity'].values[0]
                                for city in city_demand['City']
                                for mode in mode_data['Mode']]) <= max_production

        #city demand constraint
        for city in city_demand['City']:
            prob += pulp.lpSum([transport_vars[(city, mode)] *
                                mode_data[mode_data['Mode'] == mode]['Capacity'].values[0]
                                    for mode in mode_data['Mode']]) <= city_demand[city_demand['City'] == city]['Doses Needed'].values[0]
    
        prob.solve()

        i=0
        current_doses=0
        available_vehicle = dict(zip(mode_data["Mode"], mode_data["Number of Vehicles"]))
        for city in city_demand['City']:
            for mode in mode_data['Mode']:
                if pulp.value(transport_vars[(city, mode)]) and pulp.value(transport_vars[(city, mode)]) > 0:
                    doses = pulp.value(transport_vars[(city, mode)]) * mode_data[mode_data['Mode'] == mode]['Capacity'].values[0]
                    cost = distance_data[distance_data['Destination'] == city]['Distance'].values[0]*pulp.value(transport_vars[(city, mode)])*mode_data[mode_data['Mode'] == mode]['Cost per KM'].values[0]
                    results.append({
                        'Day': day,
                        'City': city,
                        'Mode': mode,
                        'Vehicles': pulp.value(transport_vars[(city, mode)]),
                        'Doses Delivered': doses,
                        'Distance': distance_data[distance_data['Destination'] == city]['Distance'].values[0], 
                        'Total Cost': cost
                    })
                    available_vehicle[mode] = available_vehicle[mode] - 1
                    summision = summision + doses
                    current_doses = current_doses + doses
                    updated_df.at[i, 'Total_Vaccinated'] = updated_df.at[i, 'Total_Vaccinated']+doses
                    updated_df.at[i, 'Uninfected'] = updated_df.at[i, 'Uninfected']-doses
            i = i + 1

        priority_map = {}  
        needed_map = {}
        remaining = max_production - current_doses
        for index,row in city_demand.iterrows():
            city = row['City']
            infection_rate = row['Infection rate']
            doses_need = updated_df.at[index, 'Uninfected']
            needed_map[city] = int(doses_need)
            if(doses_need != 0):
                priority_map[city] = int(rate_weight[infection_rate] * doses_need)
            else:
                priority_map[city] = 0
        priority_map = dict(sorted(priority_map.items(), key=lambda item: item[1], reverse=True))
        vehicle_cap = dict(zip(mode_data["Mode"], mode_data["Capacity"]))

        if(remaining>0 and sum(available_vehicle.values())>0):
            for key, value in priority_map.items():
                if(value>0 and sum(available_vehicle.values())>0 and remaining>0):
                    mindif = 0
                    vehi_mode = ""
                    for key2, value in vehicle_cap.items():
                        dif = abs(value-remaining)
                        if((dif<mindif) or (mindif==0)):
                            mindif = dif
                            vehi_mode = key2
                    if(len(vehi_mode)>0):
                        available_vehicle[vehi_mode] = available_vehicle[vehi_mode] - 1
                        delivered = min(needed_map[key],remaining)
                        remaining = remaining - delivered
                        summision = summision + delivered
                        cost = distance_data[distance_data['Destination'] == key]['Distance'].values[0]*mode_data[mode_data['Mode'] == vehi_mode]['Cost per KM'].values[0]
                        results.append({
                            'Day': day,
                            'City': key,
                            'Mode': vehi_mode,
                            'Vehicles': 1,
                            'Doses Delivered': delivered,
                            'Distance': distance_data[distance_data['Destination'] == key]['Distance'].values[0],
                            'Total Cost': cost
                        })
                        for index,row in updated_df.iterrows():
                            if(row['City']==key):
                                updated_df.at[index, 'Total_Vaccinated'] = updated_df.at[index, 'Total_Vaccinated']+delivered
                                updated_df.at[index, 'Uninfected'] = updated_df.at[index, 'Uninfected']-delivered
                else:
                    break

        day = day + 1

        for index,row in city_demand.iterrows():
            positive = int(updated_df.at[index, 'Uninfected'] * rate_weight[row['Infection rate']])
            city_demand.at[index, 'Doses Needed'] = updated_df.at[index, 'Uninfected'] - positive
            updated_df.at[index,'Uninfected'] =  updated_df.at[index, 'Uninfected'] - positive
            updated_df.at[index,'Infected'] =  updated_df.at[index, 'Infected'] + positive
    
    last_day = int(results[-1]["Day"])
    for index,row in city_demand.iterrows():
        city_demand.at[index, 'Doses Needed'] = updated_df.at[index, 'Infected']

    #starting vaccinated infected people
    summision2 = 0
    custom2 = 0

    while(sum(updated_df['Infected']>0)):
        for index,row in updated_df.iterrows():
            if(updated_df.at[index,'Infected']<mincapacity):
                custom2 += updated_df.at[index,'Infected']
                updated_df.at[index,'Infected'] = 0

        prob = pulp.LpProblem("Vaccine_Delivery2", pulp.LpMaximize)

        #dicition var for Transport_(city, mode)
        transport_vars = pulp.LpVariable.dicts("Transport",
                                            ((city, mode) for city in city_demand['City'] 
                                            for mode in mode_data['Mode']),
                                            lowBound=0, cat='Integer')

        #objective function
        prob += pulp.lpSum([transport_vars[(city, mode)] * (1/mode_data[mode_data['Mode'] == mode]['Cost per dose per KM'].values[0])
                                for city in city_demand['City'] 
                                for mode in mode_data['Mode']])
    
        #vehicle constraint
        for mode in mode_data['Mode']:
            prob += pulp.lpSum([transport_vars[(city, mode)] 
                                for city in city_demand['City']]) <= mode_data[mode_data['Mode'] == mode]['Number of Vehicles'].values[0]

        #max tranportation limit
        prob += pulp.lpSum([transport_vars[(city, mode)] * mode_data[mode_data['Mode'] == mode]['Capacity'].values[0]
                                for city in city_demand['City']
                                for mode in mode_data['Mode']]) <= max_production

        #city demand constraint
        for city in city_demand['City']:
            prob += pulp.lpSum([transport_vars[(city, mode)] *
                                mode_data[mode_data['Mode'] == mode]['Capacity'].values[0]
                                    for mode in mode_data['Mode']]) <= city_demand[city_demand['City'] == city]['Doses Needed'].values[0]
    
        prob.solve()

        k=0
        current_doses=0
        available_vehicle = dict(zip(mode_data["Mode"], mode_data["Number of Vehicles"]))
        for city in city_demand['City']:
            for mode in mode_data['Mode']:
                if pulp.value(transport_vars[(city, mode)]) and pulp.value(transport_vars[(city, mode)]) > 0:
                    doses = pulp.value(transport_vars[(city, mode)]) * mode_data[mode_data['Mode'] == mode]['Capacity'].values[0]
                    cost = distance_data[distance_data['Destination'] == city]['Distance'].values[0]*pulp.value(transport_vars[(city, mode)])*mode_data[mode_data['Mode'] == mode]['Cost per KM'].values[0]
                    results.append({
                        'Day': day,
                        'City': city,
                        'Mode': mode,
                        'Vehicles': pulp.value(transport_vars[(city, mode)]),
                        'Doses Delivered': doses,
                        'Distance': distance_data[distance_data['Destination'] == city]['Distance'].values[0], 
                        'Total Cost': cost
                    })
                    available_vehicle[mode] = available_vehicle[mode] - 1
                    summision2 = summision2 + doses
                    current_doses = current_doses + doses
                    updated_df.at[k, 'Total_Vaccinated'] = updated_df.at[k, 'Total_Vaccinated']+doses
                    updated_df.at[k, 'Infected'] = updated_df.at[k, 'Infected']-doses
            k = k + 1

        needed_map = {}
        remaining = max_production - current_doses
        for index,row in city_demand.iterrows():
            city = row['City']
            doses_need = updated_df.at[index, 'Infected']
            needed_map[city] = int(doses_need)
        needed_map = dict(sorted(needed_map.items(), key=lambda item: item[1], reverse=True))
        vehicle_cap = dict(zip(mode_data["Mode"], mode_data["Capacity"]))

        if(int(remaining)>0 and sum(available_vehicle.values())>0):
            for key, value in needed_map.items():
                if((value>0) and (sum(available_vehicle.values())>0) and (remaining>0)):
                    mindif = 0
                    vehi_mode = ""
                    for key2, value2 in vehicle_cap.items():
                        dif = abs(value2-remaining)
                        if(((dif<mindif) or (mindif==0)) and available_vehicle[key2]>0):
                            mindif = dif
                            vehi_mode = key2
                    if(len(vehi_mode)):
                        available_vehicle[vehi_mode] = available_vehicle[vehi_mode] - 1
                        delivered = min(needed_map[key],remaining)
                        remaining = remaining - delivered
                        summision2 = summision2 + delivered
                        cost = distance_data[distance_data['Destination'] == key]['Distance'].values[0]*mode_data[mode_data['Mode'] == vehi_mode]['Cost per KM'].values[0]
                        results.append({
                            'Day': day,
                            'City': key,
                            'Mode': vehi_mode,
                            'Vehicles': 1,
                            'Doses Delivered': delivered,
                            'Distance': distance_data[distance_data['Destination'] == key]['Distance'].values[0],
                            'Total Cost': cost
                        })
                        for index,row in updated_df.iterrows():
                            if(row['City']==key):
                                updated_df.at[index, 'Total_Vaccinated'] = updated_df.at[index, 'Total_Vaccinated']+delivered
                                updated_df.at[index, 'Infected'] = updated_df.at[index, 'Infected']-delivered
                else:
                    break

        day = day + 1

        for index,row in city_demand.iterrows():
            city_demand.at[index, 'Doses Needed'] = updated_df.at[index, 'Infected']

    # Final output
    return pd.DataFrame(results),last_day,custom,custom2,summision,summision2

# Run optimization
results_df,last_day2,custom,custom2,summision,summision2 = run_optimization()

# Dash Application
app = Dash(__name__,external_stylesheets=[dbc.themes.FLATLY])

total_vaccines_delivered = 112000
infected_people = 60000
uninfected_people = 52000
uninfected = summision
infected = summision2
custom_total = custom+custom2

vehicle_count = {
    "Truck":0,
    "Train":0,
    "Airplane":0
}
vehicle_cost = {
    "Truck":0,
    "Train":0,
    "Airplane":0
}
for index,row in results_df.iterrows():
    vehicle_count[row["Mode"]] = vehicle_count[row["Mode"]] + 1
    vehicle_cost[row["Mode"]] = row["Total Cost"] + vehicle_cost[row["Mode"]]

app.layout =  dbc.Container([
    html.H1("💉 Vaccination Optimization Model", className="text-center my-4"),
    
    html.H2("Input Data"),
    html.Div([
        html.Div([
            html.H5(f"{file}"),
            dash_table.DataTable(
                data=df.to_dict('records'),
                columns=[{"name": i, "id": i} for i in df.columns],
                style_table={'overflowX': 'auto'},
                style_cell={
                    'textAlign': 'center',
                    'padding': '8px',
                    'fontFamily': 'Arial',
                    'fontSize': '14px'
                },
                style_header={
                    'backgroundColor': 'rgb(200, 200, 200)',
                    'fontWeight': 'bold',
                    'textAlign': 'center'
                },
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
        }],
        style_header={
            'backgroundColor': 'rgb(200, 200, 200)',
            'fontWeight': 'bold',
            'textAlign': 'center'
        },
        style_cell={
            'textAlign': 'center',
            'padding': '8px',
            'fontFamily': 'Arial',
            'fontSize': '14px'
        },
    ),

    dbc.Card([
        dbc.CardBody([
            dbc.Row([
                dbc.Col([
                    html.H4("Total Cost", className="card-title"),
                    html.H4(f"{sum(results_df['Total Cost'])}", className="text-dark")
                ]),
                dbc.Col([
                    html.H4("Total Number Of Days", className="card-title"),
                    html.H4(f"{results_df.iloc[-1]["Day"]}", className="text-secondary")
                ]),
                dbc.Col([
                    html.H4("Total Vaccine Produced", className="card-title"),
                    html.H4(f"{(results_df.iloc[-1]["Day"])*max_production}", className="text-success")
                ]),
            ], className="mb-4"),
            dbc.Row([
                dbc.Col([
                    html.H4("Delivery By Truck", className="card-title"),
                    html.H4(f"{vehicle_count["Truck"]}", className="text-warning")
                ]),
                dbc.Col([
                    html.H4("Delivery By Train", className="card-title"),
                    html.H4(f"{vehicle_count["Train"]}", className="text-info")
                ]),
                dbc.Col([
                    html.H4("Delivery By Plane", className="card-title"),
                    html.H4(f"{vehicle_count["Airplane"]}", className="text-primary")
                ])
            ], className="mb-4"),
            dbc.Row([
                dbc.Col([
                    html.H4("Total Cost Of The Truck", className="card-title"),
                    html.H4(f"{vehicle_cost["Truck"]}", className="text-warning")
                ]),
                dbc.Col([
                    html.H4("Total Cost Of The Train", className="card-title"),
                    html.H4(f"{vehicle_cost["Train"]}", className="text-info")
                ]),
                dbc.Col([
                    html.H4("Total Cost Of The Plane", className="card-title"),
                    html.H4(f"{vehicle_cost["Airplane"]}", className="text-primary")
                ])
            ], className="mb-4"),
            dbc.Row([
                dbc.Col([
                    html.H4("Total Days To Vaccinate: UnInfected", className="card-title"),
                    html.H4(f"{last_day2}", className="text-success")
                ]),
                dbc.Col([
                    html.H4("Total Days To Vaccinate: Infected", className="card-title"),
                    html.H4(f"{results_df.iloc[-1]["Day"]-last_day2}", className="text-danger")
                ]),
                dbc.Col([
                    html.H4("Total = UnInfected+Infected+Manual", className="card-title"),
                    html.H4(f"{total} = {uninfected}+{infected}+{custom_total}", className="text-primary")
                ])
            ], className="mb-4"),
        ])
    ], className="shadow p-4 mb-5 bg-white rounded mt-3")
],fluid = True)

if __name__ == '__main__':
    app.run(debug=True)