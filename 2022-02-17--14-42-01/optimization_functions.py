    import numpy as np
from scipy import signal
from scipy import optimize
# ************** function definitions ****************
# ************************************************ AV *********************************************
def ModelDev1_L1(CalCoeff, *args): # Arg list given by scipy.optimize
	# Calculate L1 norm of difference between model values and reference values 
    # This is used as objective function for optimisation
	# CalCoeff is matrix of callibration coefficients for the model. o, O, P, Q 
    # *args is a list of 6 arrays; 3 vectors of reference values and 3 vectors sensor data (x,y,z)
    
    F = args[0] # reference data 
    M = args[1] # sensor data 
    # Buid cal coef matrix (concat o, O, P, Q)
    rows = 4
    oO = np.empty((rows,3))
    oO[:,0] = CalCoeff[0:4] # Coef for x
    oO[:,1] = CalCoeff[4:8] # Coef for y
    oO[:,2] = CalCoeff[8:12] # Coef for z

    # Calculate model values
    ModelValues = np.matmul(M, oO)

	# Calculate difference between reference and model values
    MdlDiff = F-ModelValues

    # Calculate L1 norm
    rows = len(F[:,0])
    norm = 0
    for i in range(rows):
        norm += np.sqrt(MdlDiff[i,0]**2 + MdlDiff[i,1]**2 + MdlDiff[i,2]**2) # sqrt(x^2+y^2+z^2)
    
    return norm

def acc_cal_1_simplex(AX, AY, AZ, TA, AXRef, AYRef, AZRef):
    print("Running simplex")
    rows = len(AX)
    M = np.empty((rows,4))
    MT = np.empty((4, rows))
    MTM = np.empty((4, 4))
    invMTM = np.empty((4, 4))
    invMTM_MT = np.empty((4, rows))
    F = np.empty((rows,3))
    OT = np.empty((4, 3))
    O = np.empty((3, 3))
    invO = np.empty((3, 3))
    
    M[:,0] = 1
    M[:,1] = AX
    M[:,2] = AY
    M[:,3] = AZ
      
    F[:,0] = AXRef
    F[:,1] = AYRef
    F[:,2] = AZRef
 
    MT = np.transpose(M)
    MTM = np.matmul(MT, M)
    invMTM = np.linalg.inv(MTM)
    invMTM_MT = np.matmul(invMTM, MT)
    OT = np.matmul(invMTM_MT, F)
    xfactors = OT[:, 0]
    yfactors = OT[:, 1]
    zfactors = OT[:, 2]

    # ****************************
    # Print L1 norm of MMSE solution
    print('L1 existing 1st order:', ModelDev1_L1(np.concatenate([xfactors, yfactors, zfactors]), F, M))

    # Build initial guess 1D vector
    StartValues = np.concatenate([xfactors, yfactors, zfactors])  
    

    # Run optimisation
    res = optimize.minimize(ModelDev1_L1, StartValues, args=(F,M), method='Nelder-Mead')
    # res = optimize.minimize(ModelDev1_L1, StartValues, args=(F,M), method='Powell')
    # res = optimize.minimize(ModelDev1_L1, StartValues, args=(F,M))
    # Print L1 norm of simplex solution
    xfactors = res.x[0:4] 
    yfactors = res.x[4:8]
    zfactors = res.x[8:12]
    print('L1 simplex 1st order:', ModelDev1_L1(np.concatenate([xfactors, yfactors, zfactors]), F, M))
    
    # Store result in same variable as original method
    OT[:,0] = xfactors
    OT[:, 1] = yfactors
    OT[:, 2] = zfactors
    # ****************************
    
    O = OT[1:4,:] # exclude offset vector, scaling coef
    O = np.transpose(O)

    invO = np.linalg.inv(O) # invert scaling matrix
    offset = OT[0,:] # extract offset vector
    offsetT = np.transpose(offset) 
    offset = -np.matmul(invO, offsetT)
    xoffset = offset[0]
    yoffset = offset[1]
    zoffset = offset[2] 
        
    T = np.average(TA)
    
    return xfactors, yfactors, zfactors, xoffset, yoffset, zoffset, T

def ModelDev3_L1(CalCoeff, *args): # Arg list given by scipy.optimize
	# Calculate L1 norm of difference between model values and reference values 
    # This is used as objective function for optimisation
	# CalCoeff is matrix of callibration coefficients for the model. o, O, P, Q 
    # *args is a list of 6 arrays; 3 vectors of reference values and 3 vectors sensor data (x,y,z)
    
    AXRef = args[0] # x-reference data 
    AYRef = args[1] # y-reference data 
    AZRef = args[2] # z-reference data 
    AX = args[3] # x-sensor data 
    AY = args[4] # y-sensor data 
    AZ = args[5] # z-sensor data 

    # Buid cal coef matrix (concat o, O, P, Q)
    if len(CalCoeff) == 30:
        rows = 10
        oOPQ = np.empty((rows,3))
        oOPQ[:,0] = CalCoeff[0:10] # Coef for x
        oOPQ[:,1] = CalCoeff[10:20] # Coef for y
        oOPQ[:,2] = CalCoeff[20:30] # Coef for z
        # Buid measurement matrix
        rows = len(AX)
        M = np.empty((rows,10))
        M[:,0] = 1
        M[:,1] = AX
        M[:,2] = AY
        M[:,3] = AZ
        M[:,4] = AX**2
        M[:,5] = AY**2
        M[:,6] = AZ**2
        M[:,7] = AX**3
        M[:,8] = AY**3
        M[:,9] = AZ**3
    else:
        rows = int(len(CalCoeff)/3)
        oOPQ = np.empty((rows,3))
        oOPQ[:,0] = CalCoeff[0:rows] # Coef for x
        oOPQ[:,1] = CalCoeff[rows:2*rows] # Coef for y
        oOPQ[:,2] = CalCoeff[2*rows:3*rows] # Coef for z
        # Buid measurement matrix
        rows = len(AX)
        M = np.empty((rows,7))
        M[:,0] = 1
        M[:,1] = AX
        M[:,2] = AY
        M[:,3] = AZ
        M[:,4] = AX**2
        M[:,5] = AY**2
        M[:,6] = AZ**2
        
    # Buid reference matrix
    rows = len(AX)
    F = np.empty((rows,3))
    F[:,0] = AXRef
    F[:,1] = AYRef
    F[:,2] = AZRef

    # Calculate model values
    ModelValues = np.matmul(M, oOPQ)

	# Calculate difference between reference and model values
    MdlDiff = F-ModelValues

    # Calculate L1 norm
    norm = 0
    for i in range(rows):
        norm += np.sqrt(MdlDiff[i,0]**2 + MdlDiff[i,1]**2 + MdlDiff[i,2]**2) # sqrt(x^2+y^2+z^2)
    
    return norm
    
def acc_cal3_simplex(AX, AY, AZ, TA, AXRef, AYRef, AZRef, Xinit, Yinit, Zinit):
    # Optimise calibration coefficients for 3rd order model with simplex method
    # StartValues is the initial guess for the calibration coefficients, must be 1D array
    
    # Build initial guess 1D vector
    StartValues = np.concatenate([Xinit, Yinit, Zinit])  
    
    # Build list of parameters to be passed to objective function
    # ArgList = [AXRef, AYRef, AZRef,AX, AY, AZ]
    
    # Run optimisation
    res = optimize.minimize(ModelDev3_L1, StartValues, args=(AXRef, AYRef, AZRef,AX, AY, AZ), method='Nelder-Mead')
    #res = optimize.minimize(ModelDev3_L1, StartValues, args=(AXRef, AYRef, AZRef,AX, AY, AZ), method='Powell')
    #res = optimize.minimize(ModelDev3_L1, StartValues, args=(AXRef, AYRef, AZRef,AX, AY, AZ))

    xfactors = res.x[0:10] 
    yfactors = res.x[10:20]
    zfactors = res.x[20:30]

    # xfactors = np.concatenate([res.x[0:7], np.zeros(3)]) # append if optimisation without 3rd order
    # yfactors = np.concatenate([res.x[7:14], np.zeros(3)])
    # zfactors = np.concatenate([res.x[14:21], np.zeros(3)])

    T = np.average(TA)
    
    return xfactors, yfactors, zfactors, T

def GModelDev3_L1(CalCoeff, *args): # Arg list given by scipy.optimize
    # Calculate L1 norm of difference between model values and reference values 
    # This is used as objective function for optimisation coefficients for gyro
	# CalCoeff is vector of callibration coefficients for the model. O 
    # *args is a list of 2 arrays; 1 for reference values and 1 for sensor data 
    
    F = args[0] # reference data 
    M = args[1] # sensor data 

    # Buid cal coef matrix O)
    O = np.empty((3,3))
    O[:,0] = CalCoeff[0:3] # Coef for x
    O[:,1] = CalCoeff[3:6] # Coef for y
    O[:,2] = CalCoeff[6:9] # Coef for z

    # Calculate model values
    ModelValues = np.matmul(M, O)

	# Calculate difference between reference and model values
    MdlDiff = F-ModelValues

    # Calculate L1 norm
    rows = len(F[:,0])
    norm = 0
    for i in range(rows):
        norm += np.sqrt(MdlDiff[i,0]**2 + MdlDiff[i,1]**2 + MdlDiff[i,2]**2) # sqrt(x^2+y^2+z^2)
    
    return norm

def calc_gyro_ortomatrix_simplex(GX1avXr, GY1avXr, GX2avXr, GZ2avXr, TG1avXr, TG2avXr, rotrateXr, GX1avYr, GY1avYr, GX2avYr, \
                         GZ2avYr, TG1avYr, TG2avYr, rotrateYr, GX1avZr, GY1avZr, GX2avZr, GZ2avZr, TG1avZr, TG2avZr, rotrateZr):
    
    xelements = 0
    yelements = 0
    zelements = 0
    
    xindex = np.zeros(len(GX1avXr))
    yindex = np.zeros(len(GX1avYr))
    zindex = np.zeros(len(GX1avZr))
    
    n=0
    for k in range(len(GX1avXr)):
        if abs(rotrateXr[k]) > 0.1:
            xelements = xelements + 1
            xindex[n] = k
            n = n+1
    n=0        
    for k in range(len(GY1avYr)):
        if abs(rotrateYr[k]) > 0.1:
            yelements = yelements + 1
            yindex[n] = k
            n = n+1
    
    n=0         
    for k in range(len(GZ2avZr)):
        if abs(rotrateZr[k]) > 0.1:
            zelements = zelements + 1
            zindex[n] = k
            n = n+1
    
#    print xindex
    rows = xelements + yelements + zelements 
    
    for GXoption in (0,1,2):
        M = np.empty((rows, 3))
        F = np.empty((rows, 3))
        MT = np.empty((3, rows))
        MTM = np.empty((3, 3))
        invMTM = np.empty((3, 3))
        invMTM_MT = np.empty((3, rows))
        OT = np.empty((3, 3))
        
        for k in range(xelements):                                                  # X rotation                   
            ind = int(xindex[k])   
            if GXoption == 0:
                M[k,0] = (GX1avXr[ind]+GX2avXr[ind])/2 -((GX1avXr[ind-1]+GX2avXr[ind-1])+(GX1avXr[ind+1]+GX2avXr[ind+1]))/4
            if GXoption == 1:
                M[k,0] = GX1avXr[ind] - (GX1avXr[ind-1] + GX1avXr[ind+1])/2
            if GXoption == 2:
                M[k,0] = GX2avXr[ind] - (GX2avXr[ind-1] + GX2avXr[ind+1])/2
            M[k,1] = GY1avXr[ind] - (GY1avXr[ind-1] + GY1avXr[ind+1])/2
            M[k,2] = GZ2avXr[ind] - (GZ2avXr[ind-1] + GZ2avXr[ind+1])/2
            
            F[k,0] = rotrateXr[ind]
            F[k,1] = 0
            F[k,2] = 0
        
        
                
        for k in range(yelements): 
            ind = int(yindex[k])                                          # Y rotation                        
            if GXoption == 0:
                M[k+xelements, 0] = (GX1avYr[ind]+GX2avYr[ind])/2 -((GX1avYr[ind-1]+GX2avYr[ind-1])+(GX1avYr[ind+1]+GX2avYr[ind+1]))/4
            if GXoption == 1:
                M[k+xelements, 0] = GX1avYr[ind] - (GX1avYr[ind-1] + GX1avYr[ind+1])/2
            if GXoption == 2:
                M[k+xelements, 0] = GX2avYr[ind] - (GX2avYr[ind-1] + GX2avYr[ind+1])/2
            
            M[k + xelements, 1] = GY1avYr[ind] - (GY1avYr[ind-1] + GY1avYr[ind+1])/2
            M[k + xelements, 2] = GZ2avYr[ind] - (GZ2avYr[ind-1] + GZ2avYr[ind+1])/2
            
            F[k + xelements, 0] = 0
            F[k + xelements, 1] = rotrateYr[ind]
            F[k + xelements, 2] = 0
        
        for k in range(zelements): 
            ind = int(zindex[k])                                          # Z rotation                        
            if GXoption == 0:
                M[k+xelements + yelements, 0] = (GX1avZr[ind]+GX2avZr[ind])/2 -((GX1avZr[ind-1]+GX2avZr[ind-1])+(GX1avZr[ind+1]+GX2avZr[ind+1]))/4
            if GXoption == 1:
                M[k+xelements + yelements, 0] = GX1avZr[ind] - (GX1avZr[ind-1] + GX1avZr[ind+1])/2
            if GXoption == 2:
                M[k+xelements + yelements, 0] = GX2avZr[ind] - (GX2avZr[ind-1] + GX2avZr[ind+1])/2
            
            M[k + xelements + yelements, 1] = GY1avZr[ind] - (GY1avZr[ind-1] + GY1avZr[ind+1])/2
            M[k + xelements + yelements, 2] = GZ2avZr[ind] - (GZ2avZr[ind-1] + GZ2avZr[ind+1])/2
            
            F[k + xelements + yelements, 0] = 0
            F[k + xelements + yelements, 1] = 0
            F[k + xelements + yelements, 2] = rotrateZr[ind]
        
        np.set_printoptions(precision=2, suppress = True)
        np.set_printoptions(formatter={'float': '{: 10.2f}'.format})
        print (M)
        #M[0,0] = 1.5*M[0,0] # create outlier
        np.set_printoptions(edgeitems=3,infstr='inf', linewidth=75, nanstr='nan', \
                            precision=8, suppress=False, threshold=1000, formatter=None)            # set back to default
        MT = np.transpose(M)
        MTM = np.matmul(MT, M)
        invMTM = np.linalg.inv(MTM)
        invMTM_MT = np.matmul(invMTM, MT)
        OT = np.matmul(invMTM_MT, F)
        
        # **********************************
        # Simplex-reflectio with least squares solution as start values
        print('GXoption: ', GXoption, 'Size OT: ', OT.shape)

        # L1 deviation of existing MMSE solustion
        print('L1 existing: ', GModelDev3_L1(np.concatenate([OT[:, 0], OT[:, 1], OT[:, 2]]), F,M))
        
        # Create vector with start values
        StartValues = np.concatenate([OT[:, 0], OT[:, 1], OT[:, 2]])

        # Optimistaion
        res = optimize.minimize(GModelDev3_L1, StartValues, args=(F, M), method='Nelder-Mead')

        OT[:, 0] = res.x[0:3] 
        OT[:, 1] = res.x[3:6]
        OT[:, 2] = res.x[6:9]

        # L1 deviation of simplex solustion
        print('L1 simplex: ', GModelDev3_L1(np.concatenate([OT[:, 0], OT[:, 1], OT[:, 2]]), F,M))
                
        # *************************************
        
        if GXoption == 0:
            xfactorsGX0 = OT[:, 0]
            yfactorsGX0 = OT[:, 1]
            zfactorsGX0 = OT[:, 2]
        if GXoption == 1:
            xfactorsGX1 = OT[:, 0]
            yfactorsGX1 = OT[:, 1]
            zfactorsGX1 = OT[:, 2]    
        if GXoption == 2:
            xfactorsGX2 = OT[:, 0]
            yfactorsGX2 = OT[:, 1]
            zfactorsGX2 = OT[:, 2]
            
    TG1 = (np.average(TG1avXr) + np.average(TG1avYr) + np.average(TG1avZr))/3
    TG2 = (np.average(TG2avXr) + np.average(TG2avYr) + np.average(TG2avZr))/3
       
    return xfactorsGX0, yfactorsGX0, zfactorsGX0, xfactorsGX1, yfactorsGX1, zfactorsGX1, xfactorsGX2, yfactorsGX2, zfactorsGX2, TG1, TG2




def calc_gyro_ortomatrix_3G_simplex(GX3avXr, GY3avXr, GZ3avXr, TG3avXr, rotrateXr, GX3avYr, GY3avYr, GZ3avYr, TG3avYr, \
                            rotrateYr, GX3avZr, GY3avZr, GZ3avZr, TG3avZr, rotrateZr):
    
    xelements = 0
    yelements = 0
    zelements = 0
    
    xindex = np.zeros(len(GX3avXr))
    yindex = np.zeros(len(GX3avYr))
    zindex = np.zeros(len(GX3avZr))
    
    n=0
    for k in range(len(GX3avXr)):
        if abs(rotrateXr[k]) > 0.1:
            xelements = xelements + 1
            xindex[n] = k
            n = n+1
    n=0        
    for k in range(len(GY3avYr)):
        if abs(rotrateYr[k]) > 0.1:
            yelements = yelements + 1
            yindex[n] = k
            n = n+1
    
    n=0         
    for k in range(len(GZ3avZr)):
        if abs(rotrateZr[k]) > 0.1:
            zelements = zelements + 1
            zindex[n] = k
            n = n+1
    
    rows = xelements + yelements + zelements 
    

    M = np.empty((rows, 3))
    F = np.empty((rows, 3))
    MT = np.empty((3, rows))
    MTM = np.empty((3, 3))
    invMTM = np.empty((3, 3))
    invMTM_MT = np.empty((3, rows))
    OT = np.empty((3, 3))
    
    for k in range(xelements):                                                  # X rotation                   
        ind = int(xindex[k])   
        M[k,0] = GX3avXr[ind] - (GX3avXr[ind-1] + GX3avXr[ind+1])/2
        M[k,1] = GY3avXr[ind] - (GY3avXr[ind-1] + GY3avXr[ind+1])/2
        M[k,2] = GZ3avXr[ind] - (GZ3avXr[ind-1] + GZ3avXr[ind+1])/2
        
        F[k,0] = rotrateXr[ind]
        F[k,1] = 0
        F[k,2] = 0
    
    
            
    for k in range(yelements): 
        ind = int(yindex[k])                                          # Y rotation                        

        M[k+xelements, 0] = GX3avYr[ind] - (GX3avYr[ind-1] + GX3avYr[ind+1])/2
        
        M[k + xelements, 1] = GY3avYr[ind] - (GY3avYr[ind-1] + GY3avYr[ind+1])/2
        M[k + xelements, 2] = GZ3avYr[ind] - (GZ3avYr[ind-1] + GZ3avYr[ind+1])/2
        
        F[k + xelements, 0] = 0
        F[k + xelements, 1] = rotrateYr[ind]
        F[k + xelements, 2] = 0
    
    for k in range(zelements): 
        ind = int(zindex[k])                                          # Z rotation                        

        M[k+xelements + yelements, 0] = GX3avZr[ind] - (GX3avZr[ind-1] + GX3avZr[ind+1])/2
        M[k + xelements + yelements, 1] = GY3avZr[ind] - (GY3avZr[ind-1] + GY3avZr[ind+1])/2
        M[k + xelements + yelements, 2] = GZ3avZr[ind] - (GZ3avZr[ind-1] + GZ3avZr[ind+1])/2
        
        F[k + xelements + yelements, 0] = 0
        F[k + xelements + yelements, 1] = 0
        F[k + xelements + yelements, 2] = rotrateZr[ind]
    
    np.set_printoptions(precision=2, suppress = True)
    np.set_printoptions(formatter={'float': '{: 10.2f}'.format})
    print (M)
    np.set_printoptions(edgeitems=3,infstr='inf', linewidth=75, nanstr='nan', \
                            precision=8, suppress=False, threshold=1000, formatter=None)            # set back to default
    MT = np.transpose(M)
    MTM = np.matmul(MT, M)
    invMTM = np.linalg.inv(MTM)
    invMTM_MT = np.matmul(invMTM, MT)
    OT = np.matmul(invMTM_MT, F)

     # **********************************
    # Simplex-reflectio with least squares solution as start values
    print('Size OT: ', OT.shape)

    # L1 deviation of existing MMSE solustion
    print('L1 existing: ', GModelDev3_L1(np.concatenate([OT[:, 0], OT[:, 1], OT[:, 2]]), F,M))
        
        # Create vector with start values
    StartValues = np.concatenate([OT[:, 0], OT[:, 1], OT[:, 2]])

        # Optimistaion
    res = optimize.minimize(GModelDev3_L1, StartValues, args=(F, M), method='Nelder-Mead')

    OT[:, 0] = res.x[0:3] 
    OT[:, 1] = res.x[3:6]
    OT[:, 2] = res.x[6:9]

    # L1 deviation of simplex solustion
    print('L1 simplex: ', GModelDev3_L1(np.concatenate([OT[:, 0], OT[:, 1], OT[:, 2]]), F,M))
                
        # *************************************

    xfactors = OT[:, 0]
    yfactors = OT[:, 1]
    zfactors = OT[:, 2]    

            
    TG3 = (np.average(TG3avXr) + np.average(TG3avYr) + np.average(TG3avZr))/3

       
    return xfactors, yfactors, zfactors, TG3
# ************************************************ AV *********************************************
    