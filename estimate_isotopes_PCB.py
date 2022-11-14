from math import factorial
import pandas as pd



numC = 12


c12 = 12
c13 = 13.003355
cl35 = 34.96885269
cl37 = 36.96590258
h1 = 1.007825031898
prob_Cl35 = 0.7576
prob_Cl37 = 0.2424
prob_C12 = 0.9884
prob_C13 = 0.0096

def pascal(n, k):
    n_fac = factorial(n)
    k_fac = factorial(k)
    n_k_fac = factorial((n-k))
    return (n_fac)/(k_fac*n_k_fac)


best_combos = []
def best_combo(target_mass):
    possibilities = {}
    probs = []
    for numCl in range(1,10):
        Cl = numCl
        numH = 10 - Cl
        for i in range(numC+1):
            for j in range(Cl+1):
                massC = (i*c13)+((numC-i)*c12)
                massCl = (j*cl37)+((Cl-j)*cl35)
                mass = massC + massCl + numH*h1
                if (abs(mass-target_mass) < 0.5):
                    
                    prob12 = prob_C12**(numC-i)
                    prob13 = prob_C13**i
                    coeff1 = pascal(numC, i)
                    prob_C = prob12*prob13*coeff1

                    prob35 = prob_Cl35**(Cl-j)
                    prob37 = prob_Cl37**j
                    coeff2 = pascal(Cl, j)
                    prob_Cl = prob35*prob37*coeff2

                    prob_overall = prob_C*prob_Cl

                    nums = ( i, numCl, j, mass, prob_overall)
                    possibilities[prob_overall] = nums
                    probs.append(prob_overall)
                    print("C13:" + str(i) + "  Cl37: " + str(j) + "  mass: " + str(mass) + "  probability: " + str(prob_overall))
    try:
        best = max(probs)
        
    except:
        print("No matching combination")
        return
    best = possibilities[best]

    best_combos.append(best)
   
    print("For mass " + str(target_mass) + ", the most likely combination is " + str(best[0]) + " C13 atoms and " + str(best[2]) + " Cl37 atoms.")
    return



dataframe1 = pd.read_excel(input_doc) #replace with excel document to read from

m_z = dataframe1["mass"].tolist()
m_z = [x for x in m_z if x == x] #drop nan values

for x in m_z:

    best_combo(x)

mass = [x[3] for x in best_combos]
numCl = [x[1] for x in best_combos]
Cl37 = [x[2] for x in best_combos]
c13 = [x[0] for x in best_combos]
probs = [x[4] for x in best_combos]
dict = {
    "mass":mass,
    "#Cl":numCl,
    "#Cl37": Cl37,
    "#C13":c13,
    "%":probs

}


data = pd.DataFrame(dict)
print(data)

data.to_excel(output_doc, sheet_name="probs", index=False) #replace with excel document to print to
