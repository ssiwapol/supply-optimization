import numpy as np
import pandas as pd
import dash_table
import plotly.graph_objects as go
from plotly.subplots import make_subplots

import mod

p = mod.PathFile()
configinfo = p.configinfo()


class Visualize:
    def __init__(self, user):
        self.user = user
        p.setuser(self.user)
        self.stream = False if configinfo['app']['run'] == 'local' else True
        try:
            self.file = p.loadfile(configinfo['file']['plot'])
            self.plot_status = mod.read_dict_from_worksheet(self.file, 'status', self.stream)
            self.df_route = pd.read_excel(self.file, sheet_name='route')
            self.df_supply = pd.read_excel(self.file, sheet_name='supply')
            self.df_warehouse = pd.read_excel(self.file, sheet_name='warehouse')
            self.df_destination = pd.read_excel(self.file, sheet_name='destination')
            self.df_trans = pd.read_excel(self.file, sheet_name='trans')
            self.dict_supply = dict(zip(self.df_supply['supply'], self.df_supply['supply_name']))
        except Exception:
            self.file = None

    def gen_header(self):
        df_trans = self.df_trans
        df_warehouse = self.df_warehouse
        total_rev = np.sum(df_trans['trans_rev'])
        total_fc = np.sum(df_warehouse['wh_fc_val'])
        total_vc = np.sum(df_trans['trans_vc'])
        total_netcon = total_rev - (total_fc + total_vc)
        header = {'total_netcon': "{:,.0f}".format(total_netcon),
                  'total_rev': "{:,.0f}".format(total_rev),
                  'total_vc': "{:,.0f}".format(total_vc),
                  'total_fc': "{:,.0f}".format(total_fc)}
        return header

    def gen_supply(self, click_supply):
        if click_supply is None or len(click_supply['points'][0]['hovertext'].split('<br>')) <= 1:
            supply_txt = "Supply: All"
        else:
            select_supply = click_supply['points'][0]['hovertext'].split('<br>')[0]
            supply_txt = "Supply: %s" % self.dict_supply[select_supply]
        return supply_txt

    @staticmethod
    def gen_table(df, size=5):
        return dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in df.columns],
            data=df.to_dict('records'),
            page_action="native",
            page_current=0,
            page_size=size,
        )

    def gen_supplytable(self, click_supply):
        if click_supply is None or len(click_supply['points'][0]['hovertext'].split('<br>')) <= 1:
            df_product = self.df_trans.copy()
            df_route = self.df_trans.copy()
        else:
            select_supply = click_supply['points'][0]['hovertext'].split('<br>')[0]
            df_product = self.df_trans[self.df_trans['supply']
                                       == select_supply].copy().reset_index()
            df_route = self.df_trans[self.df_trans['supply'] == select_supply].copy().reset_index()
        df_product = df_product.groupby(
            ['prod', 'prod_name'], as_index=False).agg({"trans_vol": "sum"})
        df_product = df_product.sort_values(
            by='trans_vol', ascending=False).reset_index(drop=True)
        df_routemap = df_route.groupby(['route_name', 'wh_name', 'dest_name'],
                                       as_index=False).agg({"trans_vol": "sum"})
        df_routemap = df_routemap.sort_values(
            by='trans_vol', ascending=False).reset_index(drop=True)
        return self.gen_table(df_product), self.gen_table(df_routemap)

    def plt_utilization(self):
        df_supply = self.df_supply
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        # utilization
        fig.add_trace(
            go.Bar(
                x=df_supply['supply_name'],
                y=df_supply['supply_utilize'],
                marker_color="#1f77b4",
                text=[("%.1f" % x)+' %' for x in df_supply['supply_utilize']*100],
                textposition="outside",
                hovertext=['%s<br>vol: %i<br>cap: %i<br>sku: %i' % (a, b, c, d) for a, b, c, d in zip(
                    df_supply['supply'], df_supply['supply_vol'],
                    df_supply['supply_cap'], df_supply['supply_sku'])],
                hoverinfo="text",
                name='% Utilization',
            ))
        # net contribution per ton
        fig.add_trace(
            go.Scatter(
                x=df_supply['supply_name'],
                y=df_supply['supply_netconperunit'],
                mode='lines',
                line=dict(color='red', width=1),
                hovertext=["{:,.0f}".format(x) for x in df_supply['supply_netconperunit']],
                hoverinfo='text',
                name='NetCon/Unit',
            ), secondary_y=True
        )
        # update layout
        fig.update_layout(
            title_text='**Select supply to view detail',
            yaxis=dict(tickformat=".2%"),
            showlegend=True,
            legend_orientation="h",
        )
        # update axis name
        fig.update_yaxes(title_text="%", range=[0, 3], secondary_y=False)
        fig.update_yaxes(title_text="NetCon/Unit", secondary_y=True)
        return fig

    def plt_routemap(self, click_supply):
        df_supply = self.df_supply
        df_destination = self.df_destination
        df_route = self.df_route
        sz = 5
        fig = go.Figure()
        # plot all supply
        fig.add_trace(
            go.Scattermapbox(
                lat=df_supply['supply_lat'],
                lon=df_supply['supply_long'],
                mode='markers',
                marker=go.scattermapbox.Marker(
                    size=sz, color='blue'
                ),
                text=df_supply['supply_name'],
                name="supply",
            )
        )
        # plot all destination
        fig.add_trace(go.Scattermapbox(
            lat=df_destination['dest_lat'],
            lon=df_destination['dest_long'],
            mode='markers',
            marker=go.scattermapbox.Marker(
                size=sz, color='crimson'
            ),
            text=df_destination['dest_name'],
            name="destination",
        ))

        if click_supply is None or len(click_supply['points'][0]['hovertext'].split('<br>')) <= 1:
            center_lat = list(df_destination['dest_lat'].dropna())
            center_lat.extend(list(df_supply['supply_lat'].dropna()))
            center_long = list(df_destination['dest_long'].dropna())
            center_long.extend(list(df_supply['supply_long'].dropna()))
        else:
            select_supply = click_supply['points'][0]['hovertext'].split('<br>')[0]
            df_route = df_route[df_route['supply'] == select_supply].reset_index()
            for i in range(len(df_route)):
                fig.add_trace(
                    go.Scattermapbox(
                        lat=[df_route['supply_lat'][i], df_route['dest_lat'][i]],
                        lon=[df_route['supply_long'][i], df_route['dest_long'][i]],
                        mode='lines',
                        line=dict(width=1, color='black'),
                        # opacity=float(plt_route['route_vol'][i]) / float(plt_route['route_vol'].max()),
                        name="%s --> %s: %.0f" % (df_route['supply_name'][i],
                                                  df_route['dest_name'][i], df_route['route_vol'][i]),
                        showlegend=False,
                    )
                )
            center_lat = list(df_route['dest_lat'].dropna())
            center_lat.extend(list(df_route['supply_lat'].dropna()))
            center_long = list(df_route['dest_long'].dropna())
            center_long.extend(list(df_route['supply_long'].dropna()))
        fig.update_layout(
            title_text='Route Map',
            showlegend=True,
            # legend_orientation="h",
            mapbox_style="carto-positron",
            mapbox_zoom=5,
            mapbox_center={"lat": np.mean(center_lat), "lon": np.mean(center_long)}
        )
        return fig
