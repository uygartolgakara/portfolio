# -*- coding: utf-8 -*-
"""
Produced on Wed Aug 23 13:33:20 2023
@author: KUY3IB
"""
import numpy as np
from scipy.optimize import minimize

# Define the formulas based on the given parameters
def compute_outputs(params):
    G5, G7, G9, G11, G13, G15, G17, G19, G21 = params
    
    # y1 = G5*(2+((50+G7)/50)+((1000+G17)/1000)+((1500+G19)/1500)-((70+G21)/70))
    # y2 = 14*(2+((G17+2000)/2000)+((G19+5000)/5000)+((G5+200)/200)+((G7+40)/40)-((G9+4)/4)+(abs(G11-20)/50))
    # y3 = 180*(4-(G19*(100-G21)/10/G5/600)-((G7+180)/180)-((G17+10000)/10000)+((G9+G13+10)/10))
    # y4 = 100*(2+((G7+45)/45)+((G17+4200)/4200)+((G19+8000)/8000)-((G21+10)/10))
    # y5 = 8*(3+((G21+10)/10)-((G17+6500)/6500)-((G13+3)/3)+(0 if G13 <= 0 else abs((G15+20)/50))-((G19+6000)/6000))
    # y6 = 22*(3+((G21+11)/11)-((G17+7500)/7500)-((G13+2.5)/2.5)+(0 if G13 <= 0 else abs((G15+30)/50))-((G19+5500)/5500))
    # y7 = 0.2*(2-((G17+2800)/2800)+(G21+10)/10 if G21 < 20 else (G21+7)/7)-(G19*(100-G21)/10/(G9+G13+G5)/600)-(0 if (G15 >= -35) or (G13 <= 0) else (G13+20)/20)
    # y8 = 200*(3+((G5+300)/300)-((G7+100)/100)+(0 if (G15 >= -20) or (G13 <= 0) else (abs(G13*G15)+300)/300 if G13*G15 >= 0 else (abs(G13*G15)+200)/200)+((G17+8000)/8000)-((G19+3000)/3000))
    # y9 = 20*(1.6+((G19+3000)/3000)+((G7+25)/25)+((G17+10000)/10000))
    
    y1 = G5*(2+((50+G7)/50)+((1000+G17)/1000)+((1500+G19)/1500)-((70+G21)/70))
    y2 = 14*(2+((G17+2000)/2000)+((G19+5000)/5000)+((G5+200)/200)+((G7+40)/40)-((G9+4)/4)+(abs(G11-20)/50))
    y3 = 180*(4-(G19*(100-G21)/10/G5/600)-((G7+180)/180)-((G17+10000)/10000)+((G9+G13+10)/10))
    y4 = 100*(2+((G7+45)/45)+((G17+4200)/4200)+((G19+8000)/8000)-((G21+10)/10))
    y5 = 8*(3+((G21+10)/10)-((G17+6500)/6500)-((G13+3)/3)+(abs((G15+20)/50) if G13>0 else 0)-((G19+6000)/6000))
    y6 = 22*(3+((G21+11)/11)-((G17+7500)/7500)-((G13+2.5)/2.5)+(abs((G15+30)/50) if G13>0 else 0)-((G19+5500)/5500))
    y7 = 0.2*(2-((G17+2800)/2800)+(((G21+10)/10) if G21<20 else ((G21+7)/7))-(G19*(100-G21)/10/(G9+G13+G5)/600)- (((G13+20)/20) if (G15>-35 and G13>0) else 0))
    y8 = 200*(3+((G5+300)/300)-((G7+100)/100)+ (((abs(G13*G15)+300)/300) if (G15>-20 and G13>0) else ((abs(G13*G15)+200)/200))+((G17+8000)/8000)-((G19+3000)/3000))
    y9 = 20*(1.6+((G19+3000)/3000)+((G7+25)/25)+((G17+10000)/10000))

    return [y1, y2, y3, y4, y5, y6, y7, y8, y9]

def objective(params):
    outputs = compute_outputs(params)
    return outputs[2]

def constraint_y1(params):
    return 200 - compute_outputs(params)[0]

def constraint_y1_lower(params):
    return compute_outputs(params)[0] - 200

def constraint_y1_upper(params):
    return 200 - compute_outputs(params)[0]

def constraint_y2(params):
    return 80 - compute_outputs(params)[1]

def constraint_y3(params):
    return 399 - compute_outputs(params)[2]

def constraint_y4(params):
    return 299 - compute_outputs(params)[3]

def constraint_y5(params):
    return 19 - compute_outputs(params)[4]

def constraint_y6(params):
    return 49 - compute_outputs(params)[5]

def constraint_y7(params):
    return 0.49 - compute_outputs(params)[6]

def constraint_y8(params):
    return 699 - compute_outputs(params)[7]

def constraint_y9(params):
    return 119 - compute_outputs(params)[8]

# Gather constraints
constraints = [
    {'type': 'ineq', 'fun': constraint_y1_lower},
    {'type': 'ineq', 'fun': constraint_y1_upper},
    {'type': 'ineq', 'fun': constraint_y2},
    {'type': 'ineq', 'fun': constraint_y3},
    {'type': 'ineq', 'fun': constraint_y4},
    {'type': 'ineq', 'fun': constraint_y5},
    {'type': 'ineq', 'fun': constraint_y6},
    {'type': 'ineq', 'fun': constraint_y7},
    {'type': 'ineq', 'fun': constraint_y8},
    {'type': 'ineq', 'fun': constraint_y9}
]


# Constraints for the parameters
bounds = [(10, 50), (-20, 20), (0, 5), (5, 30), (0, 5), (-100, -5), (250, 1800), (1000, 2500), (0, 50)]

# Starting point
initial_guess = [(b[0] + b[1]) / 2 for b in bounds]

# Call the minimizer
result = minimize(objective, initial_guess, bounds=bounds, constraints=constraints, method='SLSQP')

# Display the result
print("Optimal parameters:", result.x)
print("Optimal outputs:", compute_outputs(result.x))

print()

for k in result.x:
    print(k)
    
print()
    
for l in compute_outputs(result.x):
    print(l)