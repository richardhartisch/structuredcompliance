import numpy as np


# input basic measurements
# give all parameters in mm
height = float(input("Gripper height h: "))
depth = float(input("Gripper depth d: "))
wall_thickness = float(input("Gripper wall thickness: "))
cb_thickness = float(input("Gripper crossbeam thickness: "))
plane_vertical = float(input("Length of plane subjected to vertical load: "))

# input infill parameters
# give angle in degrees
# give density in decimal, i.e 10% = 0.1
angle = float(input("Infill angle phi: "))
density = float(input("Infill density rho: "))


# young's modulus in N/mm^2
young_modulus = float(input("Young's modulus E: "))

def stiffness_matrix(h,d, t,cb_t,phi, den,E):
    #Calculates the stiffness matrix of a gripper

    #h, gripper height [mm]
    #d, gripper depth [mm]
    #t, wall thickness [mm]
    #cb_t, cross beam thickness [mm]
    #theta, infil angle [deg]
    #den, infil percent, i.e. 0.1 for ten percent
    #E, Young's Modulus [N/mm**2]

    # substitute model, basic paramaters
    h_subst = h
    d_subst = d
    w_subst = (2*t + 1*cb_t)

    #plane subjected to vertical load
    A = d_subst*plane_vertical

    # compensation for
    c_i = den*10

    # Inertia of substitute model
    I = (d_subst*(w_subst**3))/12

    # factor s_i
    # phi in degrees
    s_i = phi/10 * 0.125

    # horizontal stiffness (y-direction)
    k_yy = c_i * ((12*E*I)/h_subst**3) * (1+s_i)

    # compensatory factor for vertical stiffness 1
    p = h_subst/wall_thickness

    # factor to account for density
    u_i = 1+den

    # factor to account for angle
    v_i = (1-phi/100)
    print(u_i)
    print(v_i)

    # vertical stiffness (z-direction)
    if den >= 0.2:
        k_zz = (E*A)/(h_subst*p)*u_i

    #else:
    k_zz = (E * A) / (h_subst * p) * u_i * v_i

    stiffness_matrix = np.array([[k_yy, 0], [0, k_zz]])

    return stiffness_matrix

k_mat=stiffness_matrix(height,depth,wall_thickness,cb_thickness,angle,density,young_modulus)


print("Stiffness Matrix:" + str(k_mat))
