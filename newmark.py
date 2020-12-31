# --------------------------------------------------------------------------
# name       : Newmark.py
# description: method of numerical: 
#                                  " Newmark's method "
#                                   - "constant Avg. Acc"
#                                   - "linear Acc"
# data:      : 20201221
# author     : garyhsieh
# --------------------------------------------------------------------------
import numpy as np
import math
import matplotlib.pyplot as plt

#import draw_func as dfc
import read_data as rda

# define pi value
pi = 3.1415926

# damping ratio, ksee
ksee = 0.05

# set stiffness k = 1
k = 1

# set natural period
Tn = []
Wn = []
# for i in range(0, 20, 0.1):
num = 0
for i in np.arange(0.0, 5, 0.1):
    Tn.append(i)
    if i == 0:
        Wn.append(10 ** 4)
    else:
        Wn.append(2 * pi / Tn[num])
    num = num + 1


# --------------------------------------------------------------------------
# name       : Newmark's method for " linear Acc. " 
# description: methods based on assumed variation of acceleration, such as
#              Newmark beta method. The external force is ground motion.
#              
#
# data       : 20201221
# author     : garyhsieh
# --------------------------------------------------------------------------
def newmark_linear_acc():
    print("newmark's method : linear acc")

    _delta_t = 0.005
    #_delta_t = 0.05
    #_delta_t = 0.1

    print("---------------------------------------------------------------------------------------")
    print("%12s %12s %12s" % ("Tn", "Wn", "damping ratio"))
    for i in range(0, len(Tn)):
        print("%12.4f %12.4f %12.4f" % (Tn[i], Wn[i], ksee))
    print("---------------------------------------------------------------------------------------")

    umax = []
    PSV = []
    PSA = []
    for i in range(0, len(Tn)):
        # setting initial u0, u0_prime, u0_dprime
        u_init = 0
        u_prime_init = 0
        u_dprime_init = 0

        _ui = u_init
        _ui_prime = u_prime_init

        # read acceleration and modify acc history record and calc equivalent force.
        time, up_acc, NS_acc, EW_acc  = rda.read_acc_history()
        P_up = []
        P_NS = []
        P_EW = []
        _up_acc0 = float(up_acc[0])
        _NS_acc0 = float(NS_acc[0])
        _EW_acc0 = float(EW_acc[0])
        for j in np.arange(0, len(time)):
            up_acc[j] = float(up_acc[j]) - _up_acc0
            NS_acc[j] = float(NS_acc[j]) - _NS_acc0
            EW_acc[j] = float(EW_acc[j]) - _EW_acc0
            time[j]   = float(time[j])
            
            # calculate ground motion equalent force, force = -mg
            P_up.append(- k * up_acc[j] / Wn[i] ** 2)
            P_NS.append(- k * NS_acc[j] / Wn[i] ** 2)
            P_EW.append(- k * EW_acc[j] / Wn[i] ** 2)
        

        # step 1: calc initial u0_dprime
        _ui_dprime = - EW_acc[0] - 2 * ksee * Wn[i] * _ui_prime - (Wn[i] ** 2) * _ui

        # step 2: calc a1, a2, a3
        a1 = k * ( (6 / _delta_t ** 2) / Wn[i] ** 2 + (6 * ksee / _delta_t) / Wn[i] )
        a2 = k * ( (6 / _delta_t) / Wn[i] ** 2 + (4 * ksee) / Wn[i] )
        a3 = k * ( 2 / Wn[i] ** 2 + _delta_t * ksee / Wn[i] )
        k_hat = k * ( 1 + a1 / k)

        """ debug
        print("-------------------------------------------------------------------------------------------")
        print("%12s %12s %12s %12s %12s %12s %12s" % ("Tn", "Wn", "delta t", "a1", "a2", "a3", "k_hat"))
        print("%12.4f %12.4f %12.4f %12.4f %12.4f %12.4f %12.4f" % (Tn[i], Wn[i], _delta_t, a1, a2, a3, k_hat))
        print("-------------------------------------------------------------------------------------------\n")
        #print("\n")

        # time, Pi, Pi_hat, ui, ui_prime, ui_dprime
        print("%12s %12s %12s %12s %12s %12s" % ("time", "Pi", "Pi_hat", "ui", "ui_prime", "ui_dprime"))
        print("-------------------------------------------------------------------------------------------")
        """

        # step 3: calc ui+1, ui+1_prime, ui+1_dprime
        ui = []
        ui_prime = []
        ui_dprime = []
        Pi_hat = []
       
        # define psv and psa
        wui = []
        dwui = []

        # setting initial value
        ui.append(_ui)
        ui_prime.append(_ui_prime)
        ui_dprime.append(_ui_dprime)
        Pi_hat.append(0)

        wui.append(Wn[i] * _ui)
        dwui.append( (Wn[i] ** 2) * _ui)

        # from P0 starting, calc EW of P force.
        for l in np.arange(0, len(P_EW) - 1):
            _Pi_hat_next = P_EW[l+1] + a1 * _ui + a2 * _ui_prime + a3 * _ui_dprime
            _ui_next = _Pi_hat_next / k_hat
            _ui_prime_next = (3 / _delta_t) * (_ui_next - _ui) - 2 * _ui_prime - (_delta_t / 2) * _ui_dprime
            _ui_dprime_next = (6 / _delta_t ** 2) * (_ui_next - _ui) - (6 / _delta_t) * _ui_prime - 2 * _ui_dprime

            ui.append(_ui_next)
            ui_prime.append(_ui_prime_next)
            ui_dprime.append(_ui_dprime_next)
            Pi_hat.append(_Pi_hat_next)

            wui.append(Wn[i] * _ui_next)
            dwui.append( (Wn[i] ** 2) * _ui_next )

            _ui = _ui_next
            _ui_prime = _ui_prime_next
            _ui_dprime = _ui_dprime_next
            
            """ debug
            # time, Pi, Pi_hat, ui, ui_prime, ui_dprime
            print("%12.4f %+12.4f %+12.4f %+12.4f %+12.4f %+12.4f" % \
                    (time[l], P_EW[l], Pi_hat[l], ui[l], ui_prime[l], ui_dprime[l]))
            """
        
        """ debug
        # time, Pi, Pi_hat, ui, ui_prime, ui_dprime
        print("%12.4f %+12.4f %+12.4f %+12.4f %+12.4f %+12.4f" % \
                (time[l+1], P_EW[l+1], Pi_hat[l+1], ui[l+1], ui_prime[l+1], ui_dprime[l+1]))
        """
        
        # get the max displace u(t) in acc hsitory.
        temp = map(abs, ui)
        umax.append(max(list(temp)))

        temp_wui = map(abs, wui)
        PSV.append(max(list(temp_wui)))

        temp_dwui = map(abs, dwui)
        PSA.append(max(list(temp_dwui)))


        """ debug
        plt.figure(figsize = (20, 6))
        plt.plot(time, ui, label = "u(t) - time, Tn = " + str(Tn[i]))
        #plt.plot(ufreex, ufree, label = "free, td/Tn =" + str(_tddtn), linestyle = "dotted")
        #plt.plot(ustx, ust, label = "ust, td/Tn =" + str(_tddtn), linestyle = "-.")
        #plt.plot(np.arange(0, _tdtn, _delta_t / Tn), ui, label = "Newmark's linear acc, td/Tn =" + str(_tddtn))
        plt.grid(True)
        plt.legend()
        plt.xlabel("time")
        plt.ylabel("u(t)")

        plt.show()
        #print("show data type of time: ", type(time[0]))
        """

    #print("show Tn len: ", len(Tn))
    #print("show Umax len: ", len(umax))

    #"""
    # plot ground acc - time
    plt.figure(figsize = (20, 6))
    plt.plot(time, EW_acc, label = "acc - time, delta t = " + str(_delta_t))
    plt.grid(True)
    plt.legend()
    plt.xlabel("time")
    plt.ylabel("ground acc")

    plt.show()

    # plot Umax - Tn
    plt.figure(figsize = (20, 6))
    plt.plot(Tn, umax, label = "Umax - Tn, damping ratio = " + str(ksee))
    plt.grid(True)
    plt.legend()
    plt.xlabel("Tn")
    plt.ylabel("Umax")

    plt.show()

    # PSV - Tn
    plt.figure(figsize = (20, 6))
    plt.plot(Tn, PSV, label = "PSV - Tn, damping ratio = " + str(ksee))
    plt.grid(True)
    plt.legend()
    plt.xlabel("Tn")
    plt.ylabel("PSV")

    plt.show()

    # PSA - Tn
    plt.figure(figsize = (20, 6))
    plt.plot(Tn, PSA, label = "PSA - Tn, damping ratio = " + str(ksee))
    plt.grid(True)
    plt.legend()
    plt.xlabel("Tn")
    plt.ylabel("PSA")

    plt.show()

    #"""


if __name__ == "__main__":
    newmark_linear_acc()
    
