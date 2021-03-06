import io
import datetime

from pytz import timezone
import numpy as np
import pandas as pd
import xlrd
import pyomo.environ as pyomo
from pyomo.opt import SolverFactory

import mod

p = mod.PathFile()

sheet_dict = {'supply': ['supply', 'supply_name', 'supply_lat', 'supply_long'],
              'product': ['prod', 'prod_name'],
              'route': ['route', 'route_name'],
              'warehouse': ['wh', 'wh_name', 'wh_lat', 'wh_long'],
              'destination': ['dest', 'dest_name', 'dest_lat', 'dest_long'],
              'supply_param': ['supply', 'supply_cap', 'supply_min', 'supply_max'],
              'supplyproduct_param': ['supply', 'prod', 'supplyprod_cap'],
              'logistics_param': ['supply', 'route', 'wh', 'dest', 'logis_cap', 'logis_min', 'logis_max'],
              'supplychain_param': ['supply', 'prod', 'route', 'wh', 'dest', 'sell_price', 'var_cost', 'trans_cost'],
              'warehouse_param': ['wh', 'wh_fc', 'wh_min_vol', 'wh_max_vol'],
              'demand_param': ['prod', 'dest', 'demand_vol']}
sheet_master = {'supply': 'supply',
              'product': 'prod',
              'route': 'route',
              'warehouse': 'wh',
              'destination': 'dest'}
col_str = ["supply", "supply_name", "prod", "prod_name",
           "route", "route_name", "wh", "wh_name", "dest", "dest_name"]


class Optimize:

    def __init__(self, user):
        self.user = user
        p.setuser(self.user)
        self.stream = False if p.config['app']['run'] == 'local' else True

    def validate_sheet(self, decoded, upload_status):
        input_file = io.BytesIO()
        writer = pd.ExcelWriter(input_file, engine='xlsxwriter')
        mod.write_dict_to_worksheet(upload_status, 'status', writer.book)
        # get all master data
        master_list = {}
        for sheet in sheet_master:
            try:
                df = pd.read_excel(io.BytesIO(decoded), sheet_name=sheet, dtype=str).dropna(subset=[sheet_dict[sheet][0]])
                master_list[sheet_master[sheet]] = list(df[sheet_master[sheet]].unique())
            except Exception:
                master_list[sheet_master[sheet]] = None
        col_master = [v for v in sheet_master.values()]
        # validate sheet
        status = {}
        for sheet in list(sheet_dict.keys()):
            try:
                # import all columns as string and convert to float in specific column
                df = pd.read_excel(io.BytesIO(decoded), sheet_name=sheet, dtype=str).dropna(subset=[sheet_dict[sheet][0]])
                for col in [x for x in df.columns if x not in col_str]:
                    df[col] = df[col].apply(lambda x: mod.converttofloat(x))
                df = df[[x for x in sheet_dict[sheet] if x in list(df.columns)]].reset_index(drop=True)
                # check if columns have duplicate value prior to primary key
                df_duplicate = df.groupby([x for x in col_master if x in df.columns], as_index=False).size().reset_index(name='cnt')
                df_duplicate = df_duplicate[df_duplicate['cnt'] > 1]
                # check if primary key have data in master sheet
                master_error = {}
                for col in [x for x in master_list if x in df.columns]:
                    master_error[col] = 0 if set(df[col].unique()) <= set(master_list[col]) else 1
                # summarize status
                status[sheet] = {}
                status[sheet]['column'] = 0 if set(list(df.columns)) >= set(sheet_dict[sheet]) else 1
                status[sheet]['master'] = 0 if sum([x for x in master_error.values()]) <= 0 else 1
                status[sheet]['duplicate'] = 0 if len(df_duplicate) <= 0 else 1
                status[sheet]['error'] = 0 if status[sheet]['column'] + status[sheet]['master'] + status[sheet]['duplicate'] <= 0 else 1
                df.to_excel(writer, sheet_name=sheet, index=False)
            except Exception:
                status[sheet] = {}
                status[sheet]['column'] = None
                status[sheet]['duplicate'] = None
                status[sheet]['error'] = 1
        writer.save()
        input_file.seek(0)
        p.savefile(input_file, p.config['file']['input'])
        return status

    def import_data(self):
        # read data and change data types
        df_dict = {}
        for sheet in sheet_dict.keys():
            df = pd.read_excel(p.loadfile(
                p.config['file']['input']), dtype=str, sheet_name=sheet, index=False)
            for col in [x for x in df.columns if x not in col_str]:
                df[col] = df[col].apply(lambda x: mod.converttofloat(x))
            df_dict[sheet] = df.dropna(subset=[df.columns[0]])
        # manipulate some tables
        df_dict['supplychain_param'] = df_dict['supplychain_param'].dropna()
        df_dict['supplychain_param'] = df_dict['supplychain_param'][df_dict['supplychain_param']['sell_price'] > 0]
        df_dict['supply_param']['supply_min_vol'] = df_dict['supply_param']['supply_cap'] * df_dict['supply_param']['supply_min']
        df_dict['supply_param']['supply_max_vol'] = df_dict['supply_param']['supply_cap'] * df_dict['supply_param']['supply_max']
        df_dict['logistics_param']['logis_min_vol'] = df_dict['logistics_param']['logis_cap'] * df_dict['logistics_param']['logis_min']
        df_dict['logistics_param']['logis_max_vol'] = df_dict['logistics_param']['logis_cap'] * df_dict['logistics_param']['logis_max']
        # combine all data
        df_combine = df_dict['supplychain_param'].copy()
        df_combine = pd.merge(df_combine, df_dict['supply'], on='supply', how='left')
        df_combine = pd.merge(df_combine, df_dict['product'], on='prod', how='left')
        df_combine = pd.merge(df_combine, df_dict['route'], on='route', how='left')
        df_combine = pd.merge(df_combine, df_dict['warehouse'], on='wh', how='left')
        df_combine = pd.merge(df_combine, df_dict['destination'], on='dest', how='left')
        df_combine = pd.merge(df_combine, df_dict['supply_param'], on='supply', how='left')
        df_combine = pd.merge(df_combine, df_dict['supplyproduct_param'], on=['supply', 'prod'], how='left')
        df_combine = pd.merge(df_combine, df_dict['logistics_param'], on=['supply', 'route', 'wh', 'dest'], how='left')
        df_combine = pd.merge(df_combine, df_dict['warehouse_param'], on='wh', how='left')
        df_combine = pd.merge(df_combine, df_dict['demand_param'], on=['prod', 'dest'], how='left')
        df_dict['combine'] = df_combine.copy()
        # save to self
        self.df_dict = df_dict

    def validate_feas(self):
        # create status and import data
        status = {}
        df_demand = self.df_dict['demand_param'].copy()
        df_combine = self.df_dict['combine'].copy()
        error_file = io.BytesIO()
        writer = pd.ExcelWriter(error_file, engine='xlsxwriter')

        # write sheet status to workbook
        sheet_status = mod.read_dict_from_worksheet(p.loadfile(p.config['file']['input']), 'status', self.stream)
        sheet_status['validate_datetime'] = datetime.datetime.now(timezone('Asia/Bangkok')).strftime("%Y-%m-%d %H:%M:%S")
        mod.write_dict_to_worksheet(sheet_status, 'status', writer.book)

        # validate 1- check if all demand have supply chain param
        a = df_demand.copy()
        b = df_combine[['supply', 'prod', 'route', 'wh', 'dest']].drop_duplicates()
        df_valid1 = pd.merge(a, b, on=['prod', 'dest'], how='left')
        df_valid1 = df_valid1.groupby(['prod', 'dest', 'demand_vol'], as_index=False).agg({"supply": "count"}).rename(columns={'supply': 'supplychain_param'})
        df_valid1['validate'] = df_valid1['supplychain_param'] > 0
        status['validate1'] = 1 if False in df_valid1['validate'].tolist() else 0
        df_valid1.to_excel(writer, sheet_name="validate1", index=False)

        # validate 2 - demand volume vs supplyprod cap (by product)
        a = df_combine[['prod', 'dest', 'demand_vol']].drop_duplicates().groupby(['prod']).sum().reset_index()
        b = df_combine[['supply', 'prod', 'supplyprod_cap']].drop_duplicates().groupby(['prod']).sum().reset_index()
        df_valid2 = pd.merge(a, b, on='prod', how='left')
        df_valid2['validate'] = df_valid2['supplyprod_cap'] >= df_valid2['demand_vol']
        status['validate2'] = 1 if False in df_valid2['validate'].tolist() else 0
        df_valid2.to_excel(writer, sheet_name="validate2", index=False)

        # validate 3 - demand volume vs logistics cap (by destination)
        a = df_combine[['prod', 'dest', 'demand_vol']].drop_duplicates().groupby(['dest']).sum().reset_index()
        b = df_combine[['supply', 'route', 'wh', 'dest', 'logis_max_vol']].drop_duplicates().groupby(['dest']).sum().reset_index()
        df_valid3 = pd.merge(a, b, on='dest', how='left')
        df_valid3['validate'] = df_valid3['logis_max_vol'] >= df_valid3['demand_vol']
        status['validate3'] = 1 if False in df_valid3['validate'].tolist() else 0
        df_valid3.to_excel(writer, sheet_name="validate3", index=False)

        # validate 4 - logistics cap vs supply cap (by supply)
        a = df_combine[['supply', 'supply_min_vol', 'supply_max_vol']].drop_duplicates()
        b = df_combine[['supply', 'route', 'wh', 'dest', 'logis_min_vol', 'logis_max_vol']].drop_duplicates().groupby(['supply']).sum().reset_index()
        df_valid4 = pd.merge(a, b, on='supply', how='left')
        df_valid4['validate1'] = df_valid4['supply_min_vol'] <= df_valid4['logis_max_vol']
        df_valid4['validate2'] = df_valid4['supply_max_vol'] >= df_valid4['logis_min_vol']
        status['validate4'] = 1 if False in df_valid4['validate1'].tolist() and False in df_valid4['validate2'].tolist() else 0
        df_valid4.to_excel(writer, sheet_name="validate4", index=False)

        # save file
        writer.save()
        error_file.seek(0)
        p.savefile(error_file, p.config['file']['error'])

        return status

    def optimize(self, solve_engine='cbc'):
        # create status
        status = {}
        start_time = datetime.datetime.now(timezone('Asia/Bangkok'))
        status['optimize_start_time'] = start_time.strftime("%Y-%m-%d %H:%M:%S")

        # set index for df
        df_supply = self.df_dict['supply'].set_index(['supply'])
        df_prod = self.df_dict['product'].set_index(['prod'])
        df_route = self.df_dict['route'].set_index(['route'])
        df_wh = self.df_dict['warehouse'].set_index(['wh'])
        df_dest = self.df_dict['destination'].set_index(['dest'])
        df_supply_param = self.df_dict['supply_param'].set_index(['supply'])
        df_supplyprod_param = self.df_dict['supplyproduct_param'].set_index(['supply', 'prod'])
        df_logis_param = self.df_dict['logistics_param'].set_index(['supply', 'route', 'wh', 'dest'])
        df_supplychain_param = self.df_dict['supplychain_param'].set_index(['supply', 'prod', 'route', 'wh', 'dest'])
        df_wh_param = self.df_dict['warehouse_param'].set_index(['wh'])
        df_demand_param = self.df_dict['demand_param'].set_index(['prod', 'dest'])

        # create dict
        supply_min_vol = dict(zip(df_supply_param.index, df_supply_param.supply_min_vol))
        supply_max_vol = dict(zip(df_supply_param.index, df_supply_param.supply_max_vol))
        supplyprod_cap = dict(zip(df_supplyprod_param.index, df_supplyprod_param.supplyprod_cap))
        wh_min_vol = dict(zip(df_wh_param.index, df_wh_param.wh_min_vol))
        wh_max_vol = dict(zip(df_wh_param.index, df_wh_param.wh_max_vol))
        logis_min_vol = dict(zip(df_logis_param.index, df_logis_param.logis_min_vol))
        logis_max_vol = dict(zip(df_logis_param.index, df_logis_param.logis_max_vol))
        sell_price = dict(zip(df_supplychain_param.index, df_supplychain_param.sell_price))
        var_cost = dict(zip(df_supplychain_param.index, df_supplychain_param.var_cost))
        trans_cost = dict(zip(df_supplychain_param.index, df_supplychain_param.trans_cost))
        wh_fc = dict(zip(df_wh_param.index, df_wh_param.wh_fc))
        demand_vol = dict(zip(df_demand_param.index, df_demand_param.demand_vol))

        # model
        model = pyomo.ConcreteModel()

        # define sets
        model.I = pyomo.Set(initialize=list(df_supply.index), doc='i_supply')
        model.J = pyomo.Set(initialize=list(df_prod.index), doc='i_product')
        model.K = pyomo.Set(initialize=list(df_route.index), doc='i_route')
        model.L = pyomo.Set(initialize=list(df_wh.index), doc='i_warehouse')
        model.M = pyomo.Set(initialize=list(df_dest.index), doc='i_destination')
        model.ID_TRANS = pyomo.Set(initialize=list(df_supplychain_param.index), doc='i_transportation')
        model.ID_WH = pyomo.Set(initialize=set(x[3] for x in df_supplychain_param.index), doc='i_warehouse_decision')

        # set parameters
        model.supply_min = pyomo.Param(model.I, initialize=supply_min_vol, default=0, mutable=True, doc='p_supply_min')
        model.supply_max = pyomo.Param(model.I, initialize=supply_max_vol, default=0, mutable=True, doc='p_supply_max')
        model.supplyprod_cap = pyomo.Param(model.I, model.J, initialize=supplyprod_cap, default=10000000, mutable=True, doc='p_supplyprod_cap')
        model.logis_min = pyomo.Param(model.I, model.K, model.L, model.M, initialize=logis_min_vol, default=0, mutable=True, doc='p_logistics_min')
        model.logis_max = pyomo.Param(model.I, model.K, model.L, model.M, initialize=logis_max_vol, default=1000000000, mutable=True, doc='p_logistics_max')
        model.sell_price = pyomo.Param(model.I, model.J, model.K, model.L, model.M, initialize=sell_price, default=0, mutable=True, doc='p_supplychain_selling_price')
        model.var_cost = pyomo.Param(model.I, model.J, model.K, model.L, model.M, initialize=var_cost, default=1000000000, mutable=True, doc='p_supplychain_vc')
        model.trans_cost = pyomo.Param(model.I, model.J, model.K, model.L, model.M, initialize=trans_cost, default=1000000000, mutable=True, doc='p_supplychain_transportation_cost')
        model.wh_fc = pyomo.Param(model.L, initialize=wh_fc, default=1000000000, mutable=True, doc='p_warehoues_fc')
        model.wh_min = pyomo.Param(model.L, initialize=wh_min_vol, default=0, mutable=True, doc='p_warehouse_min')
        model.wh_max = pyomo.Param(model.L, initialize=wh_max_vol, default=1000000000, mutable=True, doc='p_warehouse_max')
        model.demand_vol = pyomo.Param(model.J, model.M, initialize=demand_vol, default=0, mutable=True, doc='p_demand_value')

        # create decision variables
        model.trans_vol = pyomo.Var(model.ID_TRANS, domain=pyomo.NonNegativeReals, bounds=(0, None), doc='v_transportation_volume')
        model.wh_decision = pyomo.Var(model.ID_WH, domain=pyomo.Integers, bounds=(0, 1), doc='v_warehouse_decision')

        # constraints
        model.c = pyomo.ConstraintList(doc='constraints')
        idr = {k: v for k, v in zip(['i', 'j', 'k', 'l', 'm'], range(5))}
        # supply min/max
        for i in set(x[idr['i']] for x in model.ID_TRANS):
            model.c.add(sum([model.trans_vol[x] for x in model.trans_vol if x[idr['i']] == i]) 
                        >= model.supply_min[i])
            model.c.add(sum([model.trans_vol[x] for x in model.trans_vol if x[idr['i']] == i]) 
                        <= model.supply_max[i])
        # supply product cap
        for i, j in set((x[idr['i']], x[idr['j']]) for x in model.ID_TRANS):
            model.c.add(sum([model.trans_vol[x] for x in model.trans_vol if x[idr['i']] == i and x[idr['j']] == j]) 
                        <= model.supplyprod_cap[(i, j)])
        # logistics min/max
        for i, k, l, m in set((x[idr['i']], x[idr['k']], x[idr['l']], x[idr['m']]) for x in model.ID_TRANS):
            model.c.add(sum([model.trans_vol[x] for x in model.trans_vol if x[idr['i']] == i and x[idr['k']] == k and x[idr['l']] == l and x[idr['m']] == m]) 
                        >= model.logis_min[(i, k, l, m)])
            model.c.add(sum([model.trans_vol[x] for x in model.trans_vol if x[idr['i']] == i and x[idr['k']] == k and x[idr['l']] == l and x[idr['m']] == m]) 
                        <= model.logis_max[(i, k, l, m)])
        # demand
        for j, m in set((x[idr['j']], x[idr['m']]) for x in model.ID_TRANS):
            model.c.add(sum([model.trans_vol[x] for x in model.trans_vol if x[idr['j']] == j and x[idr['m']] == m]) 
                         == model.demand_vol[(j, m)])
        # warehouse decision, min/max
        max_vol = sum([model.demand_vol[x].value for x in model.demand_vol])
        for l in set(x[idr['l']] for x in model.ID_TRANS):
            model.c.add(sum([model.trans_vol[x] for x in model.trans_vol if x[idr['l']] == l]) 
                         <= max_vol * model.wh_decision[l])
            model.c.add(sum([model.trans_vol[x] for x in model.trans_vol if x[idr['l']] == l]) 
                         >= model.wh_min[l])
            model.c.add(sum([model.trans_vol[x] for x in model.trans_vol if x[idr['l']] == l]) 
                         <= model.wh_max[l])

        # objective Function
        model.objective = pyomo.Objective(
            expr=sum((model.trans_vol[(i, j, k, l, m)]*(model.sell_price[(i, j, k, l, m)]-model.var_cost[(i, j, k, l, m)]-model.trans_cost[(i, j, k, l, m)]))
                     for i, j, k, l, m in list(model.trans_vol)) - sum((model.wh_decision[l]*model.wh_fc[l]) for l in set(x[idr['l']] for x in model.trans_vol)),
            sense=pyomo.maximize)

        # solve
        if p.config['solver'][solve_engine] == "None":
            s = SolverFactory(solve_engine)
        else:
            s = SolverFactory(solve_engine, executable=p.config['solver'][solve_engine])
        results = s.solve(model)

        # result
        status['optimize_solver_engine'] = solve_engine
        status['optimize_solver_status'] = str(results['Solver'][0]['Status'])
        status['optimize_termination_condition'] = str(results['Solver'][0]['Termination condition'])
        end_time = datetime.datetime.now(timezone('Asia/Bangkok'))
        status['optimize_end_time'] = end_time.strftime("%Y-%m-%d %H:%M:%S")
        status['optimize_solvetime_sec'] = (end_time - start_time).total_seconds()

        # export output
        data = []
        for x in model.trans_vol:
            data.append({
                "supply": x[0],
                "prod": x[1],
                "route": x[2],
                "wh": x[3],
                "dest": x[4],
                "vol": model.trans_vol[x].value
            })
        df_output = pd.DataFrame(data)
        df_output = df_output[['supply', 'prod', 'route', 'wh', 'dest', 'vol']]
        df_combine = self.df_dict['combine'].copy()
        df_output_combine = pd.merge(df_combine, df_output, on=['supply', 'prod', 'route', 'wh', 'dest'], how='left')
        self.df_dict['output'] = df_output.copy()
        self.df_dict['output_combine'] = df_output_combine.copy()
        return status

    def gen_plot(self, opt_status):
        plot_file = io.BytesIO()
        writer = pd.ExcelWriter(plot_file, engine='xlsxwriter')

        # write sheet status to workbook
        sheet_status = mod.read_dict_from_worksheet(p.loadfile(p.config['file']['input']), 'status', self.stream)
        sheet_status.update(opt_status)
        sheet_status['plot_datetime'] = datetime.datetime.now(timezone('Asia/Bangkok')).strftime("%Y-%m-%d %H:%M:%S")
        mod.write_dict_to_worksheet(sheet_status, 'status', writer.book)

        df_combine = self.df_dict['output_combine'].copy()
        # route
        df_route = df_combine.copy()
        df_route = df_route.groupby(['supply', 'supply_name', 'supply_lat', 'supply_long',
                                     'dest', 'dest_name', 'dest_lat', 'dest_long'], as_index=False).agg({'vol': 'sum'}).rename(columns={'vol': 'route_vol'})
        df_route.to_excel(writer, sheet_name='route', index=False)

        # supply
        df_supply = df_combine.copy()
        df_supply['supply_vol'] = df_supply['vol'].fillna(0)
        df_supply['supply_netcon'] = df_supply.apply(lambda x: x['supply_vol']*(x['sell_price']+x['var_cost']-x['trans_cost']), axis=1)
        df_supply = df_supply.groupby(['supply', 'supply_name', 'supply_lat', 'supply_long',
                                       'supply_cap', 'supply_min', 'supply_max',
                                       'supply_min_vol', 'supply_max_vol'], as_index=False).agg({'supply_vol': 'sum', 'supply_netcon': 'sum', 'prod': 'count'}).rename(columns={'prod': 'supply_sku'})
        df_supply['supply_utilize'] = df_supply['supply_vol'] / df_supply['supply_cap']
        df_supply['supply_netconperunit'] = df_supply['supply_netcon'] / df_supply['supply_vol']
        df_supply.to_excel(writer, sheet_name='supply', index=False)

        # warehouse
        df_wh = df_combine.copy()
        df_wh = df_wh.groupby(['wh', 'wh_name', 'wh_fc', 'wh_min_vol', 'wh_max_vol'], as_index=False).agg({'vol': 'sum'}).rename(columns={'vol': 'wh_vol'})
        df_wh['wh_fc_val'] = df_wh.apply(lambda x: x['wh_fc'] if x['wh_vol'] > 0 else 0, axis=1)
        df_wh.to_excel(writer, sheet_name='warehouse', index=False)

        # destination
        a = df_combine[['dest', 'dest_name', 'dest_lat', 'dest_long', 'prod', 'demand_vol']].drop_duplicates(
        ).groupby(['dest', 'dest_name', 'dest_lat', 'dest_long']).sum().reset_index()
        b = df_combine[['dest', 'vol']].groupby(['dest']).sum().rename(columns={'vol': 'dest_vol'}).reset_index()
        df_dest = pd.merge(a, b, on='dest', how='left')
        df_dest.to_excel(writer, sheet_name='destination', index=False)

        # transportation
        df_trans = df_combine.copy()
        df_trans['trans_vol'] = df_trans['vol']
        df_trans['trans_rev'] = df_trans['vol'] * df_trans['sell_price']
        df_trans['trans_vc'] = df_trans['vol'] * (df_trans['var_cost'] + df_trans['trans_cost'])
        df_trans['trans_netcon'] = df_trans['trans_rev'] - df_trans['trans_vc']
        df_trans = df_trans[['supply', 'supply_name', 'supply_lat', 'supply_long',
                             'prod', 'prod_name', 'route', 'route_name', 'wh', 'wh_name',
                             'dest', 'dest_name', 'dest_lat', 'dest_long',
                             'trans_vol', 'trans_rev', 'trans_vc', 'trans_netcon']]
        df_trans.to_excel(writer, sheet_name='trans', index=False)

        writer.save()
        plot_file.seek(0)
        p.savefile(plot_file, p.config['file']['plot'])

    def gen_output(self, opt_status):
        output_file = io.BytesIO()
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

        # write sheet status to workbook
        sheet_status = mod.read_dict_from_worksheet(p.loadfile(p.config['file']['input']), 'status', self.stream)
        sheet_status.update(opt_status)
        sheet_status['output_datetime'] = datetime.datetime.now(timezone('Asia/Bangkok')).strftime("%Y-%m-%d %H:%M:%S")

        # transportation
        df_trans = self.df_dict['output_combine'].copy()
        df_trans['trans_vol'] = df_trans['vol']
        df_trans['trans_rev'] = df_trans['vol'] * df_trans['sell_price']
        df_trans['trans_vc'] = df_trans['vol'] * (df_trans['var_cost'] + df_trans['trans_cost'])
        df_trans['trans_netcon'] = df_trans['trans_rev'] - df_trans['trans_vc']
        df_trans = df_trans[['supply', 'supply_name', 'prod', 'prod_name', 'route', 'route_name', 
                             'wh', 'wh_name', 'dest', 'dest_name', 
                             'sell_price', 'var_cost', 'trans_cost',
                             'trans_vol', 'trans_rev', 'trans_vc', 'trans_netcon']]

        # supply
        a = self.df_dict['supply']
        b = self.df_dict['supply_param']
        c = self.df_dict['output_combine'][['supply', 'prod',
                                            'sell_price', 'var_cost', 'trans_cost', 'vol']]
        df_supply = pd.merge(a, b, on='supply', how='left')
        df_supply = pd.merge(df_supply, c, on='supply', how='left')
        df_supply['supply_vol'] = df_supply['vol'].fillna(0)
        df_supply['supply_netcon'] = df_supply.apply(lambda x: x['supply_vol'] * (x['sell_price']-x['var_cost']-x['trans_cost']), axis=1)
        df_supply = df_supply.groupby(['supply', 'supply_name', 'supply_cap',
                                       'supply_min', 'supply_max', 'supply_min_vol', 'supply_max_vol'], as_index=False).agg({'supply_vol': 'sum', 'supply_netcon': 'sum', 'prod': 'nunique'}).rename(columns={'prod': 'supply_sku'})
        df_supply['supply_utilize'] = df_supply['supply_vol'] / df_supply['supply_cap']
        df_supply['supply_utilize'] = df_supply['supply_utilize'].fillna(0)
        df_supply['supply_netconperunit'] = df_supply['supply_netcon'] / df_supply['supply_vol']
        df_supply['supply_netconperunit'] = df_supply['supply_netconperunit'].fillna(0)

        # warehouse
        a = self.df_dict['warehouse']
        b = self.df_dict['warehouse_param']
        c = self.df_dict['output_combine'][['wh', 'vol']]
        df_wh = pd.merge(a, b, on='wh', how='left')
        df_wh = pd.merge(df_wh, c, on='wh', how='left')
        df_wh['wh_vol'] = df_wh['vol'].fillna(0)
        df_wh = df_wh.groupby(['wh', 'wh_name', 'wh_fc', 'wh_min_vol', 'wh_max_vol'], as_index=False).agg({'wh_vol': 'sum'})
        df_wh['wh_fc_val'] = df_wh.apply(lambda x: x['wh_fc'] if x['wh_vol'] > 0 else 0, axis=1)

        # destination
        a = self.df_dict['destination']
        b = self.df_dict['demand_param']
        c = self.df_dict['output_combine'][['prod', 'dest', 'vol']]
        df_dest = pd.merge(a, b, on='dest', how='left')
        df_dest = df_dest.groupby(['dest', 'dest_name'], as_index=False).agg({'demand_vol': 'sum'})
        df_dest = pd.merge(df_dest, c, on='dest', how='left')
        df_dest['dest_vol'] = df_dest['vol'].fillna(0)
        df_dest = df_dest.groupby(['dest', 'dest_name', 'demand_vol'], as_index=False).agg({'dest_vol': 'sum'})

        # total net contribution
        total_rev = np.sum(df_trans['trans_rev'])
        total_fc = np.sum(df_wh['wh_fc_val'])
        total_vc = np.sum(df_trans['trans_vc'])
        total_netcon = total_rev - (total_fc + total_vc)
        sum_netcon = {'Net Contribution': total_netcon,
                      'Revenue': total_rev,
                      'Variable Cost': total_vc,
                      'Fixed Cost': total_fc}

        # write sheet
        mod.write_dict_to_worksheet(sheet_status, 'status', writer.book)
        mod.write_dict_to_worksheet(sum_netcon, 'summary', writer.book)
        df_trans.to_excel(writer, sheet_name='trans', index=False)
        df_supply.to_excel(writer, sheet_name='supply', index=False)
        df_wh.to_excel(writer, sheet_name='warehouse', index=False)
        df_dest.to_excel(writer, sheet_name='destination', index=False)

        writer.save()
        output_file.seek(0)
        p.savefile(output_file, p.config['file']['output'])
