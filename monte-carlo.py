
# coding: utf-8

import pandas as pd
import numpy as np
import math
import matplotlib.pyplot as plt
from matplotlib import style
from pandas_datareader import data as web
import scipy.stats


class monte_carlo:
    
    def __init__(self, start, end):
        self.start = start
        self.end = end
        
    def get_asset(self, symbol):
        start = self.start
        end = self.end

        prices = web.get_data_yahoo(symbol, start, end)['Adj Close']
        returns = prices.pct_change()
        
        self.returns = returns
        self.prices = prices
        
    
    def monte_carlo_sim(self, num_simulations, predicted_days):
        returns = self.returns
        prices = self.prices
        
        last_price = prices[-1]
        simulation_df = pd.DataFrame()
        
        for x in range(num_simulations):
            count = 0
            avg_daily_ret = returns.mean()
            variance = returns.var()
            
            daily_vol = returns.std()
            daily_drift = avg_daily_ret - (variance / 2)
            drift = daily_drift - 0.5 * daily_vol ** 2
            
            prices = []
            
            shock = drift + daily_vol * np.random.randn()
            last_price * math.exp(shock)
            prices.append(last_price)
            
            for i in range(predicted_days):
                if count == 251:
                    break
                shock = drift + daily_vol * np.random.randn()
                price = prices[count] * math.exp(shock)
                prices.append(price)
                
                count += 1
                
            simulation_df[x] = prices
            self.simulation_df = simulation_df
            self.predicted_days = predicted_days
            
            
    def monte_carlo_obpi(self, num_simulations, predicted_days, k, rf, sigma):
        returns = self.returns
        prices = self.prices
        
        last_price = prices[-1]
        monte_df = pd.DataFrame()
        obpi_df = pd.DataFrame()
        final_df = pd.DataFrame()
        
        for x in range(num_simulations):
            count = 0
            avg_daily_ret = returns.mean()
            variance = returns.var()
            
            daily_vol = returns.std()
            daily_drift = avg_daily_ret - (variance / 2)
            drift = daily_drift - 0.5 * daily_vol ** 2
            
            prices = []
            
            shock = drift + daily_vol * np.random.randn()
            last_price * math.exp(shock)
            prices.append(last_price)
            
            for i in range(predicted_days):
                if count == 251:
                    break
                shock = drift + daily_vol * np.random.randn()
                price = prices[count] * math.exp(shock)
                prices.append(price)
                
                count += 1
            
            monte_df['x'] = prices
            
            Nd1 = []
            Nd2 = []
            tt = []
            W1 = []
            W2 = []
            ss = []
            t = 1
            count = 0
            for j in range(predicted_days):
                if count == 251:
                    break
                s = monte_df['x'][count]
                d1 = math.log(s/k) + (rf + 0.5 * sigma **2) * t
                d1 = d1/(sigma * math.sqrt(t))
                nd1 = scipy.stats.norm(0, 1).cdf(d1)
                md2 = -(d1 - sigma * math.sqrt(t))
                mnd2 = scipy.stats.norm(0, 1).cdf(md2)
                w1 = (s * nd1) / (s * nd1 + k * math.exp(-rf * t) * mnd2)
                w2 = 1 - w1
                W1.append(w1)
                W2.append(w2)
                Nd1.append(nd1)
                Nd2.append(mnd2)
                tt.append(t)
                ss.append(s)
                t = t - 1/predicted_days
                count += 1
                
            obpi_df['s'] = ss
            obpi_df['w1'] = W1
            obpi_df['w2'] = W2
            obpi_df['Nd1'] = Nd1
            obpi_df['Nd2'] = Nd2
            
            V = []
            v = 1
            t = 1
            count = 0
            for l in range(predicted_days):
                if count == 199:
                    break
                v1 = (v * obpi_df['w1'][l]/obpi_df['s'][l]) * obpi_df['s'][l+1]
                v2 = v * obpi_df['w2'][l] * math.exp(-rf * t)
                v = v1 + v2
                t = t - 1/predicted_days
                V.append(v)
                count += 1
                
            final_df[x] = V     
            self.final_df = final_df
            self.predicted_days = predicted_days    
        
    def line_graph(self):
        prices = self.prices
        predicted_days = self.predicted_days
        simulation_df = self.simulation_df

        last_price = prices[-1]
        fig = plt.figure()
        style.use('bmh')
        
        title = 'Monte Carlo Simulation: '+ str(predicted_days)+ 'Days'
        plt.plot(simulation_df)
        fig.suptitle(title, fontsize=18, fontweight='bold')
        plt.xlabel('Days')
        plt.ylabel('Price ($USD)')
        plt.grid(True, color='gray')
        plt.show()
        
    def line_graph1(self):
        prices = self.prices
        predicted_days = self.predicted_days
        final_df = self.final_df
        
        last_price = prices[-1]
        fig = plt.figure()
        style.use('bmh')
        
        title = 'Monte Carlo OBPI: '+ str(predicted_days)+ 'Days'
        plt.plot(final_df)
        fig.suptitle(title, fontsize=18, fontweight='bold')
        plt.xlabel('Days')
        plt.ylabel('Percentage (%)')
        plt.grid(True, color='gray')
        plt.show()



