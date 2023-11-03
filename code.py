import random
import math
import time
import sys
sys.path.append(r"C:/Program Files (x86)/CST Studio Suite 2022/AMD64/python_cst_libraries")
import cst
import cst.interface
#print(cst.__file__)
from cst.interface import DesignEnvironment


#fitness function
thresholdSAR = 1
def fitness_function(phases):
    de = DesignEnvironment.connect_to_any_or_new()
    cst_file = r"C:\\Users\\LENOVO\\Desktop\\old_radius_old_tumor_oldring.cst"

    updateparameter = f"""
    Sub Main 
    With Solver 
         .ResetExcitationModes 
         .SParameterPortExcitation "False" 
         .SimultaneousExcitation "True" 
         .SetSimultaneousExcitAutoLabel "False" 
         .SetSimultaneousExcitationLabel "Simulation_mian" 
         .SetSimultaneousExcitationOffset "Phaseshift" 
         .PhaseRefFrequency "5.8" 
         .ExcitationSelectionShowAdditionalSettings "False" 
         .ExcitationPortMode "1", "1", "{phases[12]}", {phases[0]}, "default", "True" 
         .ExcitationPortMode "2", "1", "{phases[13]}", {phases[1]}, "default", "True" 
         .ExcitationPortMode "3", "1", "{phases[14]}", {phases[2]}, "default", "True" 
         .ExcitationPortMode "4", "1", "{phases[15]}", {phases[3]}, "default", "True" 
         .ExcitationPortMode "5", "1", "{phases[16]}", {phases[4]}, "default", "True" 
         .ExcitationPortMode "6", "1", "{phases[17]}", {phases[5]}, "default", "True" 
         .ExcitationPortMode "7", "1", "{phases[18]}", {phases[6]}, "default", "True" 
         .ExcitationPortMode "8", "1", "{phases[19]}", {phases[7]}, "default", "True" 
         .ExcitationPortMode "9", "1", "{phases[20]}", {phases[8]}, "default", "True" 
         .ExcitationPortMode "10", "1", "{phases[21]}", {phases[9]}, "default", "True" 
         .ExcitationPortMode "11", "1", "{phases[22]}", {phases[10]}, "default", "True" 
         .ExcitationPortMode "12", "1", "{phases[23]}", {phases[11]}, "default", "True" 
    End With

    End Sub
    """

    extractSAR = """
    Option Explicit
    Sub Main

        With SAR
            .Reset
            .PowerlossMonitor "loss (f=5.8) [Simulation_mian]"
            .AverageWeight (1)
            .SetOption "CST C95.3"
            .SetOption "rescale 1.0"
            .SetOption "scale accepted"
            .Create
        End With

    End Sub

    """

    SARloc = """
    Sub Main
    	SelectTreeItem ("2D/3D Results\SAR\SAR (f=5.8) [Simulation_mian] (1g)")
    	With ASCIIExport
        .Reset
        .FileName (".\\ascii.txt")
        .Mode ("FixedWidth")
        .StepX (1)
        .StepY (1)
        .StepZ (1)
        .Execute
    	End With
    End Sub

    """

    prj = de.open_project(cst_file)
    prj.schematic.execute_vba_code(updateparameter)
    
    prj.modeler.run_solver()
    time.sleep(3)
    prj.schematic.execute_vba_code(extractSAR)
    time.sleep(5)
    prj.schematic.execute_vba_code(SARloc)
    time.sleep(1)
    prj.close()

    #tumor location as a cube
    xmin =-10
    xmax = 10
    ymin =-10
    ymax = 10
    zmin =-10
    zmax = 10

    maxSARtumor = -1
    maxSARout = -1
    

    loc = "C:\\Users\\LENOVO\\Desktop\\old_radius_old_tumor_oldring\\Model\\3D\\ascii.txt"
    with open(loc, 'r') as file:
        lineNum = 0
        for line in file:
            if(lineNum<2):
                lineNum+=1
                continue
            strvalues = line.split()
            values = [float(i) for i in strvalues]
            if values[0]<xmax and values[0]>xmin and values[1]>ymin and values[1]<ymax and values[2]>zmin and values[2]<zmax:
                maxSARtumor = max(maxSARtumor,values[3])
            elif values[0]<32 and values[0]>-32 and values[1]>-32 and values[1]<32 and values[2]>0 and values[2]<32:
                maxSARout = max(maxSARout,values[3])
    #print(maxSARout)
    #print(maxSARtumor)

    return maxSARout/maxSARtumor


# PSO Parameters
num_particles = 10
max_iterations = 2
c1 = 1  # Cognitive parameter
c2 = 2  # Social parameter
w = 0.4   # Inertia weight


# Define the search space (for a single variable)
min_phase = 0
max_phase = 360


#phase calaculation

#given
t = [-15,0]   #tumor location = t
r = 80   #radius = r
lmd = 0.05172   #lambda
epsroot = (1)**0.5
l_eff = lmd/epsroot

#antenna locations
a = [[r*math.cos(i*math.pi/6),r*math.sin(i*math.pi/6)] for i in range(0,12)]
#print(a)

#distance to tumor
dist = [((t[0]-a[i][0])**2 + (t[1]-a[i][1])**2)**0.5 for i in range(len(a))]
#print(dist)

phase = [(360*dist[i]/l_eff) for i in range(len(dist))]

minphase = min(phase)

for i in range(len(phase)):
    phase[i] = round((phase[i] - minphase)%360,2)

li = [0]*11
li[0] = phase
ran = random.random()*72
for i in range(4):
    ar = []
    for j in range(12):
        ar.append((phase[j]+ran+72*i)%360)
    li[i+1] = ar

for i in range(5,11,2):
    k=1
    ar1=[]
    ar2=[]
    for j in range(12):
        if j<6:
            ar1.append((phase[j]+60*k)%360)
            ar2.append(phase[j])
        else:
            ar1.append(phase[j])
            ar2.append((phase[j]*60*k)%360)
    k+=1
    li[i]=ar1
    li[i+1]=ar2
for i in li:
    for j in range(len(i)):
        i[j] = round(i[j],2)
#print(li)


#Amplitude

ampmin = 1
ampmax = 5
r = ampmax-ampmin

for i in range(len(li)):
    for j in range(12):
        li[i].append(round(1+r*random.random(),2))
#print(li)




# Initialize particles
particles = [{'position': li[i],
              'velocity': [0]*24,
              'personal_best_position': li[i],
              'personal_best_value': float('inf')}
             for i in range(num_particles)]
#print(particles)

# Initialize global best
global_best_position = None
global_best_value = float('inf')


# Main PSO loop
for iteration in range(max_iterations):
    print("\n \n")
    print("gb: ",global_best_position)
    particlenum = 0
    for particle in particles:
        particlenum += 1
        # Evaluate the objective function
        value = fitness_function(particle['position'])
        print("Iteration No: ", iteration, ". Antenna",particlenum)
        

        # Update personal best if needed
        if value < particle['personal_best_value']:
            particle['personal_best_position'] = particle['position']
            particle['personal_best_value'] = value
            print("Personal best updated.",particle['position'])
        else: print("Personal best not updated. ",particle['position'])

        # Update global best if needed
        if value < global_best_value:
            global_best_position = particle['position']
            global_best_value = value
            print("Global best position updated. ",particle['position'])


        # Update particle velocity and position for each dimension
        for i in range(24):
            r1 = random.random()
            r2 = random.random()
            cognitive_velocity = c1 * r1 * (particle['personal_best_position'][i] - particle['position'][i])
            social_velocity = c2 * r2 * (global_best_position[i] - particle['position'][i])
            particle['velocity'][i] = w * particle['velocity'][i] + cognitive_velocity + social_velocity
            particle['position'][i] += particle['velocity'][i]
            if i<12:
                particle['position'][i] = particle['position'][i] % 360
                if particle['position'][i] < 0: particle['position'][i] += 360
            if i>=12:
                if particle['position'][i] < ampmin: particle['position'][i] = ampmin
                elif particle['position'][i] > ampmax: particle['position'][i] = ampmax

# Print the optimized result
print("Optimal Solution: x =", [round(xi, 2) for xi in global_best_position])
print("Optimal Value:", round(global_best_value, 2))

#the values of  tumor position and definition of tumor cube has to be updated
