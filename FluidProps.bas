'FLUIDPROPS: A library of diverse fluid properties and equations (incl. psychrometrics, thermodanamics and fluid flow)
'- All input and output parameters are in SI units
'- References are given for all equations & data. Main references are NIST Chemistry WebBook, and ASHRAE Fundamentals Handbook 2017 (based on ASHRAE LibHuAirProp)
'- The code calculates two kinds of properties differently. This approach ensures maximum consistency of dependencies between the different calculated properties:
'   (i) FUNDAMENTAL properties, based on accurate experimental data. Examples are: density, heat capacity, thermal conductivity, dynamic viscosity
'   (ii) DERIVED properties, calculated from one or more fundamental properties using formulae. Examples are kinematic viscosity [v=µ/rho], thermal diffusivity [a=k/(rho·cp)] and Prandtl number [Pr=µ·cp/k]
'  Examples of (i) fundamentalare
'- Original version is available at https://github.com/SchildCode/FluidProps
'- If you use this code, please cite it. BibTeX-format reference is given on page https://github.com/SchildCode/FluidProps

'Copyright original author: Peter.Schild@OsloMet.no, 2000-2024
'This source-code is licensed freely as Attribute-ShareAlike under CC BY-SA 4.0 (https://creativecommons.org/licenses/by-sa/4.0/).
'It is provided without warranty of any kind.
'Please do give author feedback if you have suggestions to improvements.

Option Explicit
Const ZeroC# = 273.15 '[K]
Const onePi# = 3.14159265358979

Private Function ErrorMsg#(ErrorText As String)
    'Here you can change the behaviour when out-of bounds values are calculated. Uncomment the line(s) for your preferred alternative(s).
    MsgBox ErrorText 'ALTERNATIVE 1: Popup message box.
    ErrorMsg = -999# 'ALTERNATIVE 2: Return unlikely value (as Double). This is useful when you want to return error-flag to the calling routine
    'Stop 'ALTERNATIVE 3: Comment out as needed
End Function

Private Function usrMaxDbl#(aa#, bb#)
    'Significantly faster than Worksheetfunction.Max(aa,bb)
    If aa < bb Then
        usrMaxDbl = bb
    Else
        usrMaxDbl = aa
    End If
End Function

Private Function usrMinDbl#(aa#, bb#)
    'Significantly faster than Worksheetfunction.Min(aa,bb)
    If aa < bb Then
        usrMinDbl = aa
    Else
        usrMinDbl = bb
    End If
End Function

'----------------------------
'LIQUID WATER & ICE
'----------------------------

Public Function Water_Dens#(ByVal Tw_K#)
    'Water_Dens [kg/m³] = Density of liquid water at standard atmospheric pressure (101325 Pa).
    'INPUTS: Tw_K [K] = Water temperature
    'NOTE: This is a fundamental thermophysical property, based on empirical data.

    Const Tcrit# = 647.096 '[K] Critical-point temperature
    Dim Trc# '[-] Complement of Reduced temperature = 1 - (T / T_critical)

    If Tw_K < 273.15 Or 373.15 < Tw_K Then
        Water_Dens = ErrorMsg("Tw out of range in water_Dens.  Tw = " & Tw_K & " K")
    Else
        Trc = 1 - Tw_K / Tcrit
        'Data source: Eric W. Lemmon, Mark O. McLinden and Daniel G. Friend, "Thermophysical Properties of Fluid Systems" in NIST Chemistry WebBook, NIST Standard Reference Database Number 69, Eds. P.J. Linstrom and W.G. Mallard, National Institute of Standards and Technology, Gaithersburg MD, 20899, https://doi.org/10.18434/T4D303, (retrieved March 31, 2020).
        'https://webbook.nist.gov/chemistry/fluid/
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.999999813, Max deviation = ±0.01168 kg/m³ in the range 273.16K to 373.12K
        Water_Dens = (1 + Trc ^ 2 * (-2.68259321562161 + 0.969985158312932 * Trc ^ 2)) / (1.5658241366722E-03 - 2.34210562352665E-03 * Trc)
    End If
End Function

Public Function Water_Cp#(ByVal Tw_K#)
    'Water_Cp [J/kgK] = Heat capacity of liquid water at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tw_K [K] = Water temperature
    'NOTE: This is a fundamental thermophysical property, based on empirical data.

    Const Tcrit# = 647.096 '[K] Critical-point temperature
    Dim Tr# '[-] Reduced temperature = T / T_critical

    If Tw_K < 273.15 Or 373.15 < Tw_K Then
        Water_Cp = ErrorMsg("Tw out of range in water_Cp.  Tw = " & Tw_K & " K")
    Else
        Tr = Tw_K / Tcrit
        'Data source: Eric W. Lemmon, Mark O. McLinden and Daniel G. Friend, "Thermophysical Properties of Fluid Systems" in NIST Chemistry WebBook, NIST Standard Reference Database Number 69, Eds. P.J. Linstrom and W.G. Mallard, National Institute of Standards and Technology, Gaithersburg MD, 20899, https://doi.org/10.18434/T4D303, (retrieved March 31, 2020).
        'https://webbook.nist.gov/chemistry/fluid/
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.999991, Max deviation = ±0.127 J/kgK in the range 273.16K to 373.12K
        Water_Cp = (-38527.626300554 * Tr ^ 5) / (1 + Tr * (-7.57372380642321 + Tr * (20.5726787552669 - 21.1539281976836 * Tr)))
    End If
End Function

Public Function Water_Conduct#(ByVal Tw_K#)
    'Water_Conduct [W/mK] = Thermal conductivity of liquid water and solid ice at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tw_K [K] = Water temperature
    'NOTE: This is a fundamental thermophysical property, based on empirical data.

    Const Tcrit# = 647.096 '[K] Critical-point temperature
    Dim Tr# '[-] Complement to 1 of reduced temperature = 1 - (T / T_critical)

    Select Case Tw_K
    Case Is < 228#
        Water_Conduct = ErrorMsg("Tw too low in water_Conduct.  Tw = " & Tw_K & " K")
    Case Is < ZeroC 'solid ice
        'REFERENCE: US Coast Guard
        Water_Conduct = 5.1143235511 - 0.0105864721 * Tw_K
    Case Is <= 373.15 'liquid water
        Tr = 1 - Tw_K / Tcrit
        'Data source: Eric W. Lemmon, Mark O. McLinden and Daniel G. Friend, "Thermophysical Properties of Fluid Systems" in NIST Chemistry WebBook, NIST Standard Reference Database Number 69, Eds. P.J. Linstrom and W.G. Mallard, National Institute of Standards and Technology, Gaithersburg MD, 20899, https://doi.org/10.18434/T4D303, (retrieved March 31, 2020).
        'https://webbook.nist.gov/chemistry/fluid/
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.999999986, Max deviation = ±1.43693E-05 W/mK in the range 273.16K to 373.12K
        Water_Conduct = (0.614572668031482 + Tr * Tr * Tr * (-8.39219943470412 + 10.3751706905586 * Tr)) / (1 + Tr * Tr * (-3.56844084636025 + 7.17981988639961 * Tr * Tr * Tr))
    Case Else
        Water_Conduct = ErrorMsg("Tw too high in water_Conduct.  Tw = " & Tw_K & " K")
    End Select
End Function

Public Function Water_Enth#(ByVal Tw_K#)
    'Water_Enth [J/kg] = Specific enthalpy of liquid water at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tw_K [K] = Water temperature
    'NOTE: This is strictly a derived thermophysical property, based on integrating dh = cp·dT over over the temperature range from reference state (T=0°C) to T_K. This is complicated by the fact that cp is itself a non-linear function of temperature. Therefore a dedicated correlation is given here

    Dim TC# '[°C]

    If Tw_K < 273.15 Or 373.15 < Tw_K Then
        Water_Enth = ErrorMsg("Tw out of range in water_Enth.  Tw = " & Tw_K & " K")
    Else
        TC = Tw_K - ZeroC
        'Data source: Eric W. Lemmon, Mark O. McLinden and Daniel G. Friend, "Thermophysical Properties of Fluid Systems" in NIST Chemistry WebBook, NIST Standard Reference Database Number 69, Eds. P.J. Linstrom and W.G. Mallard, National Institute of Standards and Technology, Gaithersburg MD, 20899, https://doi.org/10.18434/T4D303, (retrieved March 31, 2020).
        'https://webbook.nist.gov/chemistry/fluid/
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.9999999979, Max deviation = ±22 J/kgK in the range 273.16K to 373.12K
        Water_Enth = (83.1183363886333 + TC * (4216.04762279028 + TC * (108.878331267255 + 5.58619783034135E-05 * TC ^ 2))) / (1 + 2.61682405418377E-02 * TC)
    End If
End Function

Public Function Water_DynaVisc#(ByVal Tw_K#)
    'Water_DynaVisc [Pa·s] = Dynamic viscosity of liquid water at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tw_K [K] = Water temperature
    'NOTE: This is a fundamental thermophysical property, based on empirical data, because it directly relates to the molecular characteristics and interactions in a fluid, independent of its density

    Const Tcrit# = 647.096 '[K] Critical-point temperature
    Dim Tr# '[-] Complement to 1 of reduced temperature = 1 - (T / T_critical)

    If Tw_K < 273.15 Or 373.15 < Tw_K Then
        Water_DynaVisc = ErrorMsg("Tw out of range in water_DynaVisc.  Tw = " & Tw_K & " K")
    Else
        Tr = 1 - Tw_K / Tcrit
        'Data source: Eric W. Lemmon, Mark O. McLinden and Daniel G. Friend, "Thermophysical Properties of Fluid Systems" in NIST Chemistry WebBook, NIST Standard Reference Database Number 69, Eds. P.J. Linstrom and W.G. Mallard, National Institute of Standards and Technology, Gaithersburg MD, 20899, https://doi.org/10.18434/T4D303, (retrieved March 31, 2020).
        'https://webbook.nist.gov/chemistry/fluid/
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.9999999874, Max deviation = ±2.297E-07 Pa·s in the range 273.16K to 373.12K
        Water_DynaVisc = (1 - 3.72194823883595 * Tr * Tr * Tr) / (100556.859960968 * Tr - 399310.024486561 * Tr * Tr + 444709.674726627 * Tr * Tr * Tr - 161766.674865504 * Tr * Tr * Tr * Tr * Tr)
    End If
End Function

Public Function Water_KineVisc#(ByVal Tw_K#)
    'Water_KineVisc [m²/s] = Kinematic viscosity of liquid water at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tw_K [K] = Water temperature
    'NOTE: This is a derived thermophysical property:
    '  v = µ/rho, where v=kinematic visc. [m²/s], µ=dynamic visc. [Pa·s], rho=density [kg/m³]

    If Tw_K < 273.15 Or 373.15 < Tw_K Then
        Water_KineVisc = ErrorMsg("Tw out of range in water_KineVisc.  Tw = " & Tw_K & " K")
    Else
        Water_KineVisc = Water_DynaVisc(Tw_K) / Water_Dens(Tw_K)
    End If
End Function

Public Function Water_Pr#(ByVal Tw_K#)
    'Water_Pr [-] = Prandtl number of liquid water at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tw_K [K] = Water temperature
    'NOTE: This is a derived thermophysical property:
    '  Pr = v/a    where v=kinematic viscosity [Pa·s], a=thermal diffusivity [W/mK]. Both v and a are themselves derived properties
    '  Pr = Cp·µ/k where Cp=heat capacity [J/kgK], µ=dynamic viscosity [Pa·s], k=thermal conductivity [W/mK]. Cp, µ and k are all fundamental properties. Therefore this equation is preferred.
    
    If Tw_K < 273.15 Or 373.15 < Tw_K Then
        Water_Pr = ErrorMsg("Tw out of range in water_Pr.  Tw = " & Tw_K & " K")
    Else
        Water_Pr = Water_Cp(Tw_K) * Water_DynaVisc(Tw_K) / Water_Conduct(Tw_K)
    End If
End Function

Public Function Water_thermDiff#(ByVal Tw_K#)
    'Water_thermDiff [m²/s] = Thermal diffusivity of liquid water at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tw_K [K] = Water temperature
    'NOTE: This is a derived thermophysical property:
    '  a = k/(rho·Cp)  where a=thermal diffusivity [m²/s], k=thermal conductivity [W/(m·K)], rho=density [kg/m³], Cp=specific heat capacity [J/(kg·K)]
    
    If Tw_K < 273.15 Or 373.15 < Tw_K Then
        Water_thermDiff = ErrorMsg("Tw out of range in Water_thermDiff.  Tw = " & Tw_K & " K")
    Else
        Water_thermDiff = Water_Conduct(Tw_K) / (Water_Dens(Tw_K) * Water_Cp(Tw_K))
    End If
End Function

'-----------------------------
'WATER VAPOUR
'-----------------------------

Private Function Vapour_Pws#(ByVal Tdry_K#, Optional ice As Boolean = False)
    'Vapour_Pws [Pa] = Saturation vapour pressure, i.e. the equilibrium water vapor pressure above a flat surface
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature. If Tdry_K is dew-point, then Vapour_Pws is vapour pressure
    '- ice = TRUE if over ice when below 0°C (e.g. chilled mirror hygrometer), default is FALSE (Meteo data is always over liquid water)
    'NOTE: This is a fundamental thermophysical property, based on empirical data
    
    'DISCUSSION OF ICE CONTRA LIQUID WATER:
    'Care needs to be taken to consider the difference between saturation vapour over ice or liquid water.
    'By default this routine assumes liquid water.
    'In meteorological practice, relative humidity is always given over liquid water (bottom eqn.).
    'If the process of heating/cooling takes place fast, temporary existing of supercooled water at temperature lower than –0.01 °C is possible.
    'As a result, formulas for calculating the pressure of saturating above water as well as ice are needed (in practice between temperatures from –18°C to the triple point +0.01°C).
    
    'CHILLED MIRROR VERSUS CAPACITIVE HYGROMETERS:
    'The equation for ice is mostly of interest for frost-point measurements when using chilled mirror hygrometers below 0°C, since these
    'instruments directly measure the temperature at which a frost layer and the overlying vapour are in equilibrium.
    'Below the triple point (T < 0.01°C), either dew or frost may exist on the mirror, which implies the mirror may be measuring
    'the dew-point (Td) or the frost-point (Tf).  Supercooled conditions not uncommon with chilled mirror hygrometers, in which case the
    'mirror surface can maintain a layer of liquid water as low as -15°C for hours or days (especially above -5°C).
    'However, assuming liquid water is always on the mirror results in far larger errors. Therefore always assume frost!
    'Capacitive RH sensors (e.g. Vaisala, Rotronic) always respond to RH relative to a plane surface of liquid water.
    
    'MORE ACCURATE ALTERNATIVE EQUATIONS
    '- IAPWS R6-95 (2016): "Revised release on the IAPWS formulation 1995 for the thermodynamic properties of ordinary water substance for general and scientific use", 19 pp., http://www.iapws.org/relguide/IAPWS95-2016.pdf
    '- IAPWS R7-97 (2012): http://www.iapws.org/relguide/IF97-Rev.html
    '- IAPWS R10-06 (2009): http://www.iapws.org/relguide/Ice-2009.html
    '- IAPWS R14-08 (2011): "Revised release on the pressure along the melting and sublimation curves of ordinary water substance", 7 pp., http://www.iapws.org/relguide/MeltSub2011.pdf.
    '  See ASHRAE Fundamentals 2021 §1.8 for more details about a slightly more accurate but slower, equation
    
    If Tdry_K < 173.15 Or 375.15 < Tdry_K Then
        Vapour_Pws = ErrorMsg("Tdry_K out of range in Function Vapour_Pws (Tdry_K=" & Tdry_K & ")")
    ElseIf Tdry_K < 273.16 And ice Then
        'ASHRAE Handbook Fundamentals 2017, Eqn.(5) for ice. Validity range approx -100°C to 0°C
        'Vapour_Pws = Exp(-5674.5359 / Tdry_K + 6.3925247 - 0.009677843 * Tdry_K + 0.00000062215701 * Tdry_K ^ 2 + 2.0747825E-09 * Tdry_K ^ 3 - 9.484024E-13 * Tdry_K ^ 4 + 4.1635019 * Log(Tdry_K))
        Vapour_Pws = Exp(-5674.5359 / Tdry_K + 6.3925247 + Tdry_K * (-0.009677843 + Tdry_K * (0.00000062215701 + Tdry_K * (2.0747825E-09 - Tdry_K * 9.484024E-13))) + 4.1635019 * Log(Tdry_K))
    Else
        'ASHRAE Handbook Fundamentals 2017, Eqn.(6) over liquid water. Validity range approx -50°C to +102°C
        'Vapour_Pws = Exp(-5800.2206 / Tdry_K + 1.3914993 - 0.048640239 * Tdry_K + 0.000041764768 * Tdry_K ^ 2 - 0.000000014452093 * Tdry_K ^ 3 + 6.5459673 * Log(Tdry_K))
        Vapour_Pws = Exp(-5800.2206 / Tdry_K + 1.3914993 + Tdry_K * (-0.048640239 + Tdry_K * (0.000041764768 - Tdry_K * 0.000000014452093)) + 6.5459673 * Log(Tdry_K))
    End If
End Function

Public Function Vapour_Cp#(ByVal Tdry_K#)
    'Vapour_Cp [J/kgK] = Heat capacity of water vapour, valid near standard atmospheric pressure (101325 Pa)
    'INPUTS: Tdry_K [K] = Dry-bulb air temperature
    'NOTE: This is a fundamental thermophysical property, based on empirical data
    
    Dim Tdry_C#

    Tdry_C = Tdry_K - ZeroC
    'REFERENCE: Reid et al. (1987 p.656f., 668)
    Vapour_Cp = 1858 + 0.382 * Tdry_C + 0.000422 * Tdry_C ^ 2 - 0.0000001996 * Tdry_C ^ 3
End Function

Public Function Vapour_Dv#(ByVal Tdry_K#, ByVal Patm_Pa#)
    'Vapour_Dv [m²/s] = Vapour mass diffusivity of water vapour in air. A.k.a. diffustion coefficient. Used to calculate Schmidt number.
    'Typical value: Dv = 25.8 mm²/s = 2.58E-5 m²/s at 25°C and standard atmospheric pressure (101325 Pa)
    'INPUTS:
    '- Tdry_K [K] = Dry-bulb air temperature
    '- Patm_Pa [Pa] = Atmospheric pressure
    'NOTE: This is a fundamental thermophysical property, based on empirical data
    
    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 6.2, Eqn.(10): Citation of Sherwood & Pigford, 1952. Valid up to 1100°C
    Vapour_Dv# = 0.000000926 * (Tdry_K ^ 2.5) / (Patm_Pa * (Tdry_K + 245)) '[m²/s]
End Function

'-----------------------------
'ANTIFREEZE COOLANT LIQUIDS
'-----------------------------

Public Function EthyleneGlycol_VolFraction#(Tfreeze_C#)
    'EthyleneGlycol_VolFraction [-] = Minimum required volume fraction of Ethylene Glycol for freezing point Tfreeze_C, valid at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tfreeze_C [°C] = Minimum freezing point
    'NOTE: This is a fundamental thermophysical property, based on empirical data
    
    If Tfreeze_C < -48.3 Or 0 < Tfreeze_C Then
        EthyleneGlycol_VolFraction = ErrorMsg("Tfreeze_C out of range in EthyleneGlycol_VolFraction.  Tfreeze_C = " & Tfreeze_C & " K")
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017, Page 31.6
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.999917452; Max error = ±0.00439850988 [Volume fraction] in the range -48°C to 0°C
        EthyleneGlycol_VolFraction = (Tfreeze_C * (-2.97952987932674E-02 + 1.42763189691829E-04 * Tfreeze_C)) / (1 - 4.31395029603815E-02 * Tfreeze_C)
    End If
End Function

Public Function EthyleneGlycol_Dens#(T_C#, VolFraction#)
    'EthyleneGlycol_Dens [kg/m³] = Density of aqueous solution of ethylene glycol coolant/antifreeze
    'INPUTS:
    '- T_C [°C] = Temperature of aqueous solution.
    '- VolFraction [-] = Volume fraction of ethylene glycol, in the range 0 to 1.
    'NOTE: You can first find the minimum required value of VolFraction with function EthyleneGlycol_VolFraction(T_C)

    If T_C < -35 Or 100 < T_C Then
        EthyleneGlycol_Dens = ErrorMsg("T_C out of range in EthyleneGlycol_Dens.  T_C = " & T_C & " °C")
    ElseIf VolFraction < 0 Or 0.9 < VolFraction Then
        EthyleneGlycol_Dens = ErrorMsg("VolFraction out of range in EthyleneGlycol_Dens.  VolFraction = " & VolFraction)
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017, Section 31
        'Fitted by symbolic regression to a polynomial by Eureqa (https://www.nutonian.com/products/eureqa/)
        'Statistics: R²=0.99997026; Max error = ±0.69203097 kg/m³ (±0.06%) in the range -35°C to 100°C
        EthyleneGlycol_Dens = 1001.1409640659 + 179.444904879415 * VolFraction - 0.182158011573945 * T_C - 0.314100479180544 * T_C * VolFraction - 2.44011693274942E-03 * T_C ^ 2 - 38.9624042140087 * VolFraction ^ 2
    End If
End Function

Public Function EthyleneGlycol_Cp#(T_C#, VolFraction#)
    'EthyleneGlycol_Cp [J/kgK] = Specific heat capacity of aqueous solution of ethylene glycol coolant/antifreeze
    'INPUTS:
    '- T_C [°C] = Temperature of aqueous solution.
    '- VolFraction [-] = Volume fraction of ethylene glycol, in the range 0 to 1.
    'NOTE: You can first find the minimum required value of VolFraction with function EthyleneGlycol_VolFraction(T_C)
    
    If T_C < -35 Or 100 < T_C Then
        EthyleneGlycol_Cp = ErrorMsg("T_C out of range in EthyleneGlycol_Cp.  T_C = " & T_C & " °C")
    ElseIf VolFraction < 0 Or 0.9 < VolFraction Then
        EthyleneGlycol_Cp = ErrorMsg("VolFraction out of range in EthyleneGlycol_Cp.  VolFraction = " & VolFraction)
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017, Section 31
        'Fitted by symbolic regression to a polynomial by Eureqa (https://www.nutonian.com/products/eureqa/)
        'Statistics: R²=0.9999987; Max error = ±1.4920783 J/kgK (±0.037%) in the range -35°C to 100°C
        EthyleneGlycol_Cp = 4097.79502411204 + 1.19750661201195 * T_C + 5.61124881968268 * T_C * VolFraction - 1557.90516584712 * VolFraction - 461.188951269786 * VolFraction ^ 2 - 0.558537199166186 * T_C * VolFraction ^ 2
    End If
End Function

Public Function EthyleneGlycol_Conduct#(T_C#, VolFraction#)
    'EthyleneGlycol_Conduct [W/mK] = Thermal conductivity of aqueous solution of ethylene glycol coolant/antifreeze
    'INPUTS:
    '- T_C [°C] = Temperature of aqueous solution.
    '- VolFraction [-] = Volume fraction of ethylene glycol, in the range 0 to 1.
    'NOTE: You can first find the minimum required value of VolFraction with function EthyleneGlycol_VolFraction(T_C)
    
    If T_C < -35 Or 100 < T_C Then
        EthyleneGlycol_Conduct = ErrorMsg("T_C out of range in EthyleneGlycol_Conduct.  T_C = " & T_C & " °C")
    ElseIf VolFraction < 0 Or 0.9 < VolFraction Then
        EthyleneGlycol_Conduct = ErrorMsg("VolFraction out of range in EthyleneGlycol_Conduct.  VolFraction = " & VolFraction)
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017, Section 31
        'Fitted by symbolic regression to a polynomial by Eureqa (https://www.nutonian.com/products/eureqa/)
        'Statistics: R²=0.99995783; Max error = ±0.00253363 W/mK (±0.8%) in the range -35°C to 100°C
        EthyleneGlycol_Conduct = 0.562795658206869 + 1.89231603003848E-03 * T_C + 0.206393536467352 * VolFraction ^ 2 + 8.32696279315093E-04 * T_C * VolFraction ^ 2 + 7.24337888092844E-06 * VolFraction * T_C ^ 2 - 0.524076791259904 * VolFraction - 2.61949570275199E-03 * T_C * VolFraction - 7.14058837551305E-06 * T_C ^ 2
    End If
End Function

Public Function EthyleneGlycol_DynaVisc#(T_C#, VolFraction#)
    'EthyleneGlycol_DynaVisc [Pa·s] = Dynamic viscosity of aqueous solution of ethylene glycol coolant/antifreeze
    'INPUTS:
    '- T_C [°C] = Temperature of aqueous solution.
    '- VolFraction [-] = Volume fraction of ethylene glycol, in the range 0 to 1.
    'NOTE: You can first find the minimum required value of VolFraction with function EthyleneGlycol_VolFraction(T_C)
    
    Dim Tr# 'reciprocal temperature deviation

    If T_C < -35 Or 100 < T_C Then
        EthyleneGlycol_DynaVisc = ErrorMsg("T_C out of range in EthyleneGlycol_DynaVisc.  T_C = " & T_C & " °C")
    ElseIf VolFraction < 0 Or 0.9 < VolFraction Then
        EthyleneGlycol_DynaVisc = ErrorMsg("VolFraction out of range in EthyleneGlycol_Dynavisc.  VolFraction = " & VolFraction)
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017, Section 31
        'Fitted by symbolic regression with Eureqa (https://www.nutonian.com/products/eureqa/)
        'Statistics: R²=0.99950351; Max error = ±0.19262573 Pa·s (Mean ±2.36%) in the range -35°C to 100°C
        Tr = 1 / (113.365314457654 + T_C)
        EthyleneGlycol_DynaVisc = Exp(-28853.879497481 * VolFraction * Tr * Tr + (377.303159990392 + 849.345796155349 * VolFraction) * Tr - 9.84889759113783 - 1.73736446942109 * VolFraction)
    End If
End Function

Public Function PropyleneGlycol_VolFraction#(Tfreeze_C#)
    'PropyleneGlycol_VolFraction [-] = Minimum required volume fraction of Propylene Glycol for freezing point Tfreeze_C, valid at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tfreeze_C [°C] = Minimum freezing point
    
    If Tfreeze_C < -51.1 Or 0 < Tfreeze_C Then
        PropyleneGlycol_VolFraction = ErrorMsg("Tfreeze_C out of range in PropyleneGlycol_VolFraction.  Tfreeze_C = " & Tfreeze_C & " K")
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017, Page 31.6
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.999973663; Max error = ±0.003282686 [Volume fraction] in the range -51°C to 0°C
        PropyleneGlycol_VolFraction = (Tfreeze_C * (-2.93493229032601E-02 - 1.59432533823575E-04 * Tfreeze_C ^ 2)) / (1 + Tfreeze_C ^ 2 * (6.66645351796462E-03 - 1.46482042633864E-04 * Tfreeze_C))
    End If
End Function

Public Function PropyleneGlycol_Dens#(T_C#, VolFraction#)
    'PropyleneGlycol_Dens [kg/m³] = Density of aqueous solution of Propylene glycol coolant/antifreeze
    'INPUTS:
    '- T_C [°C] = Temperature of aqueous solution.
    '- VolFraction [-] = Volume fraction of Propylene glycol, in the range 0 to 1.
    'NOTE: You can first find the minimum required value of VolFraction with function PropyleneGlycol_VolFraction(T_C)
    If T_C < -35 Or 100 < T_C Then
        PropyleneGlycol_Dens = ErrorMsg("T_C out of range in PropyleneGlycol_Dens.  T_C = " & T_C & " °C")
    ElseIf VolFraction < 0 Or 0.9 < VolFraction Then
        PropyleneGlycol_Dens = ErrorMsg("VolFraction out of range in PropyleneGlycol_Dens.  VolFraction = " & VolFraction)
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017, Section 31
        'Fitted by symbolic regression to a polynomial by Eureqa (https://www.nutonian.com/products/eureqa/)
        'Statistics: R²=0.99885777; Max error = ±4.637788 kg/m³ (±0.49%) in the range -35°C to 100°C
        PropyleneGlycol_Dens = 1001.60967056623 + 116.278983583743 * VolFraction + 0.003052484815767 * T_C ^ 2 * VolFraction ^ 2 - 8.53826710430915E-02 * T_C - 0.737217611426134 * T_C * VolFraction - 3.19173967918996E-03 * T_C ^ 2 - 53.8301051153873 * VolFraction ^ 3
    End If
End Function

Public Function PropyleneGlycol_Cp#(T_C#, VolFraction#)
    'PropyleneGlycol_Cp [J/kgK] = Specific heat capacity of aqueous solution of Propylene glycol coolant/antifreeze
    'INPUTS:
    '- T_C [°C] = Temperature of aqueous solution.
    '- VolFraction [-] = Volume fraction of Propylene glycol, in the range 0 to 1.
    'NOTE: You can first find the minimum required value of VolFraction with function PropyleneGlycol_VolFraction(T_C)

    If T_C < -35 Or 100 < T_C Then
        PropyleneGlycol_Cp = ErrorMsg("T_C out of range in PropyleneGlycol_Cp.  T_C = " & T_C & " °C")
    ElseIf VolFraction < 0 Or 0.9 < VolFraction Then
        PropyleneGlycol_Cp = ErrorMsg("VolFraction out of range in PropyleneGlycol_Cp.  VolFraction = " & VolFraction)
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017, Section 31
        'Fitted by symbolic regression to a polynomial by Eureqa (https://www.nutonian.com/products/eureqa/)
        'Statistics: R²=0.99999523; Max error = ±2.7606174 J/kgK (±0.09%) in the range -35°C to 100°C
        PropyleneGlycol_Cp = 4131.79124057325 + 1.0587116373911 * T_C + 5.60018700825693 * T_C * VolFraction - 786.13680726358 * VolFraction - 49.8664661311868 * VolFraction ^ 6 - 1134.40440780431 * VolFraction ^ 2
    End If
End Function

Public Function PropyleneGlycol_Conduct#(T_C#, VolFraction#)
    'PropyleneGlycol_Conduct [W/mK] = Thermal conductivity of aqueous solution of Propylene glycol coolant/antifreeze
    'INPUTS:
    '- T_C [°C] = Temperature of aqueous solution.
    '- VolFraction [-] = Volume fraction of Propylene glycol, in the range 0 to 1.
    'NOTE: You can first find the minimum required value of VolFraction with function PropyleneGlycol_VolFraction(T_C)
    
    If T_C < -35 Or 100 < T_C Then
        PropyleneGlycol_Conduct = ErrorMsg("T_C out of range in PropyleneGlycol_Conduct.  T_C = " & T_C & " °C")
    ElseIf VolFraction < 0 Or 0.9 < VolFraction Then
        PropyleneGlycol_Conduct = ErrorMsg("VolFraction out of range in PropyleneGlycol_Conduct.  VolFraction = " & VolFraction)
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017, Section 31
        'Fitted by symbolic regression to a polynomial by Eureqa (https://www.nutonian.com/products/eureqa/)
        'Statistics: R²=0.99993905; Max error = ±0.00277531 W/mK (±1%) in the range -35°C to 100°C
        PropyleneGlycol_Conduct = 0.566116039484301 + 1.9173284065513E-03 * T_C + 0.20750768738397 * VolFraction ^ 2 + 7.59544959739093E-04 * T_C * VolFraction ^ 2 + 7.25403057090419E-06 * VolFraction * T_C ^ 2 - 0.582804237084253 * VolFraction - 2.63636971222475E-03 * T_C * VolFraction - 7.29003289330946E-06 * T_C ^ 2
    End If
End Function

Public Function PropyleneGlycol_DynaVisc#(T_C#, VolFraction#)
    'PropyleneGlycol_DynaVisc [Pa·s] = Dynamic viscosity of aqueous solution of Propylene glycol coolant/antifreeze
    'INPUTS:
    '- T_C [°C] = Temperature of aqueous solution.
    '- VolFraction [-] = Volume fraction of Propylene glycol, in the range 0 to 1.
    'NOTE: You can first find the minimum required value of VolFraction with function PropyleneGlycol_VolFraction(T_C)
    
    Dim Tr# 'reciprocal temperature deviation

    If T_C < -35 Or 100 < T_C Then
        PropyleneGlycol_DynaVisc = ErrorMsg("T_C out of range in PropyleneGlycol_DynaVisc.  T_C = " & T_C & " °C")
    ElseIf VolFraction < 0 Or 0.9 < VolFraction Then
        PropyleneGlycol_DynaVisc = ErrorMsg("VolFraction out of range in PropyleneGlycol_DynaVisc.  VolFraction = " & VolFraction)
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017, Section 31
        'Fitted by symbolic regression with Eureqa (https://www.nutonian.com/products/eureqa/)
        'Statistics: R²=0.99925065; Max error = ±0.27060081 Pa·s (Mean ±3.7%) in the range -35°C to 100°C
        Tr = 1 / (216.452232513663 + T_C)
        PropyleneGlycol_DynaVisc = Exp(391867.744949403 * Tr * Tr + (1834.88325375033 * VolFraction - 1809.46299784498) * Tr - 6.32340689214633 - 3.83110457708013 * VolFraction)
    End If
End Function

'-----------------------------
'REFRIGERANTS
'-----------------------------

Function Refrigerant_Sbubble#(Refrigerant$, TK#)
    'Refrigerant_Sbubble [kJ/kg] = Bubble-point entropy, i.e. left side of Ts-diagram
    'INPUTS:
    '- Refrigerant [string] = Refrigerant ASHRAE-code e.g. "R410a"
    '- TK [K] = Refrigerant temperature at bubble-point [K]
    'SOURCE: Fitted to tables in ASHRAE Handbook Fundamentals, 2017, Chapter 30
    'NOTE: Equations generated by symbolic regression with https://github.com/SchildCode/RatFun-Regression:
    '  Average statistics for all 9 refrigerants: R² = 0.9992, Max Absolute Error = ±0.038 kJ/kg over the whole temperature range
    'AUTHOR: P.G.Schild, 2020

    Dim Tr# 'Complement to 1 of reducted temperature, i.e. 1 - (T [K] / Tcritical [K]), where Tcritical is the refrigerant's critical point
    Dim Sc# 'Entropy at refrigerant's critical point [kJ/(kgK)]

    Select Case Refrigerant
    Case "R32" 'CAS# 75-10-5: HFC "Difluoromethane", https://echa.europa.eu/substance-information/-/substanceinfo/100.000.764
        Tr = usrMaxDbl(0, 1 - TK / 351.26)
        Sc = 1.6486
        Refrigerant_Sbubble = Sc - (Tr * (15.1860219519526 + Tr * (143.148871606057 + 154.369735731194 * Tr * Tr * Tr))) / (1 + Tr * (68.0718307045745 + 2.08783990667309 * Tr))
    Case "R134a" 'CAS# 811-97-2: HFC "Norflurane" aka 1,1,1,2-Tetrafluoroethane, https://echa.europa.eu/substance-information/-/substanceinfo/100.011.252
        Tr = usrMaxDbl(0, 1 - TK / 374.2)
        Sc = 1.5621
        Refrigerant_Sbubble = Sc - (Tr * (9.66688452454042 + Tr * (124.51193401481 + 87.882560248025 * Tr * Tr * Tr))) / (1 + Tr * (75.6962662137555 - 5.77561006753399 * Tr))
    Case "R290" 'CAS# 74-98-6: HC Propane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.753
        Tr = usrMaxDbl(0, 1 - TK / 369.89)
        Sc = 2.0516
        Refrigerant_Sbubble = Sc - (Tr * (8.56145481255502 + Tr * (35.510705198475 + 18.5210905201964 * Tr * Tr * Tr))) / (1 + Tr * (14.2493896051139 - 3.94295854870169 * Tr))
    Case "R407c" 'Zerotropic HFC blend R-32+125+134a (23±2/25±2/52±2%), https://en.wikipedia.org/wiki/R-407C
        Tr = usrMaxDbl(0, 1 - TK / 359.18)
        Sc = 1.5384
        Refrigerant_Sbubble = Sc - (Tr * (12.0937091475015 + Tr * (144.068346027271 + 132.520960498277 * Tr * Tr * Tr))) / (1 + Tr * (83.4894309710863 - 1.27877450209623 * Tr))
    Case "R410a" 'Zerotropic HFC blend R-32+125 (50+.5,–1.5/50+1.5,–.5%), https://en.wikipedia.org/wiki/R-410A
        Tr = usrMaxDbl(0, 1 - TK / 344.51)
        Sc = 1.5181
        Refrigerant_Sbubble = Sc - (Tr * (11.8920967219117 + Tr * (105.378762642689 + 92.8988132330928 * Tr * Tr * Tr))) / (1 + Tr * (61.7062405514335 - 5.41121070779363 * Tr))
    Case "R600" 'CAS# 106-97-8: Butane, https://echa.europa.eu/substance-information/-/substanceinfo/100.003.136
        Tr = usrMaxDbl(0, 1 - TK / 425.13)
        Sc = 2.3631
        Refrigerant_Sbubble = Sc - (Tr * (11.7531025603875 + Tr * (146.664843481129 + 96.5147138011199 * Tr * Tr * Tr))) / (1 + Tr * (44.9953777537236 + 1.32317305669742 * Tr))
    Case "R600a" 'CAS# 75-28-5: Isobutane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.780
        Tr = usrMaxDbl(0, 1 - TK / 407.81)
        Sc = 2.2259
        Refrigerant_Sbubble = Sc - (Tr * (16.6771635516396 + Tr * (269.516471103558 + 154.474379472207 * Tr * Tr * Tr))) / (1 + Tr * (83.3067470338159 + 4.12393369005376 * Tr))
    Case "R717" 'CAS# 7664-41-7: Ammonia (NH3), https://echa.europa.eu/substance-information/-/substanceinfo/100.028.760
        Tr = usrMaxDbl(0, 1 - TK / 405.4)
        Sc = 3.5542
        Refrigerant_Sbubble = Sc - (Tr * (46.6217766617561 + Tr * (320.380840016196 + 146.290765084784 * Tr * Tr * Tr))) / (1 + Tr * (62.3383436370725 - 17.1613521295301 * Tr))
    Case "R744" 'CAS# 124-38-9: Carbon dioxide (CO2), https://echa.europa.eu/substance-information/-/substanceinfo/100.004.271
        Tr = usrMaxDbl(0, 1 - TK / 304.13)
        Sc = 1.4336
        Refrigerant_Sbubble = Sc - (Tr * (27.1966097899245 + Tr * (361.985272011556 + 889.007724088429 * Tr * Tr * Tr))) / (1 + Tr * (133.382664055875 + 48.0011035103798 * Tr))
    Case Else
        MsgBox "Refrigerant '" & Refrigerant & "' is not in database", vbExclamation
    End Select
End Function

Function Refrigerant_Sdew#(Refrigerant$, TK#)
    'Refrigerant_Sdew [kJ/kgK] = Dew-point entropy, i.e. right side of Ts-diagram
    'INPUTS:
    '- Refrigerant [string] = Refrigerant ASHRAE-code e.g. "R410a"
    '- TK [K] = Refrigerant temperature at dew-point [K]
    'SOURCE: Fitted to tables in ASHRAE Handbook Fundamentals, 2017, Chapter 30
    'NOTE: Equations generated by symbolic regression with https://github.com/SchildCode/RatFun-Regression:
    '  Average statistics for all 9 refrigerants: R² = 0.9992, Max Absolute Error = ±0.038 kJ/kg over the whole temperature range
    'AUTHOR: P.G.Schild, 2020

    Dim Tr# 'Complement to 1 of reducted temperature, i.e. 1 - (T [K] / Tcritical [K]), where Tcritical is the refrigerant's critical point
    Dim Sc# 'Entropy at refrigerant's critical point [kJ/(kgK)]

    Select Case Refrigerant
    Case "R32" 'CAS# 75-10-5: HFC "Difluoromethane", https://echa.europa.eu/substance-information/-/substanceinfo/100.000.764
        Tr = usrMaxDbl(0, 1 - TK / 351.26)
        Sc = 1.6486
        Refrigerant_Sdew = Sc - (Tr * (-13.242688036753 + Tr * (-7.85942757928493 - 12.9119465504887 * Tr * Tr * Tr))) / (1 + Tr * (33.7237924262883 - 38.0678948048016 * Tr))
    Case "R134a" 'CAS# 811-97-2: HFC "Norflurane" aka 1,1,1,2-Tetrafluoroethane, https://echa.europa.eu/substance-information/-/substanceinfo/100.011.252
        Tr = usrMaxDbl(0, 1 - TK / 374.2)
        Sc = 1.5621
        Refrigerant_Sdew = Sc - (Tr * (-9.51052079751144 + Tr * (7.21511712934949 - 84.2188159972129 * Tr * Tr * Tr))) / (1 + Tr * (58.6621646913184 - 51.2420967661778 * Tr))
    Case "R290" 'CAS# 74-98-6: HC Propane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.753
        Tr = usrMaxDbl(0, 1 - TK / 369.89)
        Sc = 2.0516
        Refrigerant_Sdew = Sc - (Tr * (-20.6942057756752 + Tr * (12.70036134274 - 164.688891778152 * Tr * Tr * Tr))) / (1 + Tr * (68.5173216390091 - 60.9109869548097 * Tr))
    Case "R407c" 'Zerotropic HFC blend R-32+125+134a (23±2/25±2/52±2%), https://en.wikipedia.org/wiki/R-407C
        Tr = usrMaxDbl(0, 1 - TK / 359.18)
        Sc = 1.5384
        Refrigerant_Sdew = Sc - (Tr * (-10.6594912900386 + Tr * (2.99671314125426 - 39.689391865346 * Tr * Tr * Tr))) / (1 + Tr * (53.9104351779721 - 65.7745697998255 * Tr))
    Case "R410a" 'Zerotropic HFC blend R-32+125 (50+.5,–1.5/50+1.5,–.5%), https://en.wikipedia.org/wiki/R-410A
        Tr = usrMaxDbl(0, 1 - TK / 344.51)
        Sc = 1.5181
        Refrigerant_Sdew = Sc - (Tr * (-9.06033100087089 + Tr * (-1.98308659288065 - 21.3907718089308 * Tr * Tr * Tr))) / (1 + Tr * (36.5331745192676 - 42.3686508350456 * Tr))
    Case "R600" 'CAS# 106-97-8: Butane, https://echa.europa.eu/substance-information/-/substanceinfo/100.003.136
        Tr = usrMaxDbl(0, 1 - TK / 425.13)
        Sc = 2.3631
        Refrigerant_Sdew = Sc - (Tr * (-15.2969975847954 + Tr * (45.5856679195242 - 227.716934021835 * Tr * Tr * Tr))) / (1 + Tr * (70.1864060308476 - 42.1559614082148 * Tr))
    Case "R600a" 'CAS# 75-28-5: Isobutane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.780
        Tr = usrMaxDbl(0, 1 - TK / 407.81)
        Sc = 2.2259
        Refrigerant_Sdew = Sc - (Tr * (-16.3952941348538 + Tr * (44.8062389226689 - 253.17952957514 * Tr * Tr * Tr))) / (1 + Tr * (76.2734611172978 - 44.2643027129568 * Tr))
    Case "R717" 'CAS# 7664-41-7: Ammonia (NH3), https://echa.europa.eu/substance-information/-/substanceinfo/100.028.760
        Tr = usrMaxDbl(0, 1 - TK / 405.4)
        Sc = 3.5542
        Refrigerant_Sdew = Sc - (Tr * (-31.7292833707121 + Tr * (2.12038637392883 + 125.751612531618 * Tr * Tr * Tr))) / (1 + Tr * (23.4298351350206 - 37.2902799144488 * Tr))
    Case "R744" 'CAS# 124-38-9: Carbon dioxide (CO2), https://echa.europa.eu/substance-information/-/substanceinfo/100.004.271
        Tr = usrMaxDbl(0, 1 - TK / 304.13)
        Sc = 1.4336
        Refrigerant_Sdew = Sc - (Tr * (-33.5633898829662 + Tr * (-319.312986430669 - 1897.38052550397 * Tr * Tr * Tr))) / (1 + Tr * (128.631833292984 + 222.971101627653 * Tr))
    Case Else
        MsgBox "Refrigerant '" & Refrigerant & "' is not in database", vbExclamation
    End Select
End Function

Function Refrigerant_CpDew#(Refrigerant$, TK#)
    'Refrigerant_CpDew [kJ/(kgK)] = Refrigerant gas-phase specific heat capacity, at dew-point temperature and pressure
    'INPUTS:
    '- Refrigerant [string] = Refrigerant ASHRAE-code e.g. "R410a"
    '- TK [K] = Refrigerant temperature at dew-point [K]
    'SOURCE: Fitted to tables in ASHRAE Handbook Fundamentals, 2017, Chapter 30
    'NOTE: Equations generated by symbolic regression with https://github.com/SchildCode/RatFun-Regression:
    '  Average statistics for all 9 refrigerants: R² =0.99997, Max Absolute Error = ±0.1196 kJ/kg over the whole temperature range
    'AUTHOR: P.G.Schild, 2020

    Dim Tr# 'Complement to 1 of reducted temperature, i.e. 1 - (T [K] / Tcritical [K]), where Tcritical is the refrigerant's critical point

    Select Case Refrigerant
    Case "R32" 'CAS# 75-10-5: HFC "Difluoromethane", https://echa.europa.eu/substance-information/-/substanceinfo/100.000.764
        Tr = usrMaxDbl(0, 1 - TK / 351.26)
        Refrigerant_CpDew = (1 + 6.92945151413515 * Tr) / (7.95296515498074E-04 + Tr * (7.56906798797785 + Tr * Tr * (41.4119987064445 - 43.6398246025174 * Tr)))
    Case "R134a" 'CAS# 811-97-2: HFC "Norflurane" aka 1,1,1,2-Tetrafluoroethane, https://echa.europa.eu/substance-information/-/substanceinfo/100.011.252
        Tr = usrMaxDbl(0, 1 - TK / 374.21)
        Refrigerant_CpDew = (1 + 11.097341049689 * Tr) / (6.955416937568E-04 + Tr * (14.1851844493013 + Tr * Tr * (36.4167084397886 - 18.6464966625328 * Tr)))
    Case "R290" 'CAS# 74-98-6: HC Propane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.753
        Tr = usrMaxDbl(0, 1 - TK / 369.89)
        Refrigerant_CpDew = (1 + 8.81128897439505 * Tr) / (-1.71291057253671E-03 + Tr * (6.52239839563111 + Tr * Tr * (13.0520585115035 - 7.71319590930056 * Tr)))
    Case "R407c" 'Zerotropic HFC blend R-32+125+134a (23±2/25±2/52±2%), https://en.wikipedia.org/wiki/R-407C
        Tr = usrMaxDbl(0, 1 - TK / 359.18)
        Refrigerant_CpDew = (1 + 9.63114541187629 * Tr) / (5.71296826196862E-02 + Tr * (12.3247804143034 + Tr * Tr * (39.8238751285889 - 29.611210148231 * Tr)))
    Case "R410a" 'Zerotropic HFC blend R-32+125 (50+.5,–1.5/50+1.5,–.5%), https://en.wikipedia.org/wiki/R-410A
        Tr = usrMaxDbl(0, 1 - TK / 344.51)
        Refrigerant_CpDew = (1 + 7.66354857414616 * Tr) / (-1.14940244991465E-02 + Tr * (10.0247292919728 + Tr * Tr * (25.8576838505004 - 7.36543568651091 * Tr)))
    Case "R600" 'CAS# 106-97-8: Butane, https://echa.europa.eu/substance-information/-/substanceinfo/100.003.136
        Tr = usrMaxDbl(0, 1 - TK / 425.13)
        Refrigerant_CpDew = (1 + 17.1386947248888 * Tr) / (-1.15384435270888E-03 + Tr * (9.34151907884644 + Tr * Tr * (29.2760562576776 - 20.9240025543195 * Tr)))
    Case "R600a" 'CAS# 75-28-5: Isobutane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.780
        Tr = usrMaxDbl(0, 1 - TK / 407.81)
        Refrigerant_CpDew = (1 + 14.803472846493 * Tr) / (4.07517047136446E-04 + Tr * (8.82622325226665 + Tr * Tr * (23.3969428536691 - 10.3769449798891 * Tr)))
    Case "R717" 'CAS# 7664-41-7: Ammonia (NH3), https://echa.europa.eu/substance-information/-/substanceinfo/100.028.760
        Tr = usrMaxDbl(0, 1 - TK / 405.4)
        Refrigerant_CpDew = (1 + 4.82962793918416 * Tr) / (1.49143706578461E-03 + Tr * (2.15038151957287 + Tr * Tr * (12.967159534515 - 16.9104094739299 * Tr)))
    Case "R744" 'CAS# 124-38-9: Carbon dioxide (CO2), https://echa.europa.eu/substance-information/-/substanceinfo/100.004.271
        Tr = usrMaxDbl(0, 1 - TK / 304.13)
        Refrigerant_CpDew = (1 + 2.48348359150248 * Tr) / (-3.71811874594367E-03 + Tr * (6.70751524823017 + Tr * Tr * (-9.14533639703711 + 27.1182140684234 * Tr)))
    Case Else
        MsgBox "Refrigerant '" & Refrigerant & "' is not in database", vbExclamation
    End Select
End Function

Function Refrigerant_Hbubble#(Refrigerant$, TK#)
    'Refrigerant_Hbubble [kJ/kgK] = Bubble-point enthalpy, i.e. left side of Ph-diagram
    'INPUTS:
    '- Refrigerant [string] = Refrigerant ASHRAE-code e.g. "R410a"
    '- TK [K] = Refrigerant temperature at bubble-point [K]
    'SOURCE: Fitted to tables in ASHRAE Handbook Fundamentals, 2017, Chapter 30
    'NOTE: Equations generated by symbolic regression with https://github.com/SchildCode/RatFun-Regression:
    '  Average statistics for all 9 refrigerants: R² = 0.9988, Max Absolute Error = ±12.3 kJ/kg over the whole temperature range
    'AUTHOR: P.G.Schild, 2020

    Dim Tr# 'Complement to 1 of reducted temperature, i.e. 1 - (T [K] / Tcritical [K]), where Tcritical is the refrigerant's critical point
    Dim Hc# 'Enthalpy at refrigerant's critical point [kJ/(kgK)]

    Select Case Refrigerant
    Case "R32" 'CAS# 75-10-5: HFC "Difluoromethane", https://echa.europa.eu/substance-information/-/substanceinfo/100.000.764
        Tr = usrMaxDbl(0, 1 - TK / 351.26)
        Hc = 414.15
        Refrigerant_Hbubble = Hc - (Tr * (5377.91956101175 + Tr * (46617.0249105985 + -26186.6467842203 * Tr + 18897.9434846015 * Tr * Tr))) / (1 + 63.9712636775809 * Tr)
    Case "R134a" 'CAS# 811-97-2: HFC "Norflurane" aka 1,1,1,2-Tetrafluoroethane, https://echa.europa.eu/substance-information/-/substanceinfo/100.011.252
        Tr = usrMaxDbl(0, 1 - TK / 374.21)
        Hc = 389.64
        Refrigerant_Hbubble = Hc - (Tr * (3394.72261017963 + Tr * (40801.2388080537 + -20357.687633116 * Tr + 11486.3069162471 * Tr * Tr))) / (1 + 65.6607987817449 * Tr)
    Case "R290" 'CAS# 74-98-6: HC Propane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.753
        Tr = usrMaxDbl(0, 1 - TK / 369.89)
        Hc = 555.24
        Refrigerant_Hbubble = Hc - (Tr * (5980.0523933281 + Tr * (62774.7502810252 + -35878.445484561 * Tr + 17789.013871877 * Tr * Tr))) / (1 + 53.2299251271831 * Tr)
    Case "R407c" 'Zerotropic HFC blend R-32+125+134a (23±2/25±2/52±2%), https://en.wikipedia.org/wiki/R-407C
        Tr = usrMaxDbl(0, 1 - TK / 359.18)
        Hc = 378.48
        Refrigerant_Hbubble = Hc - (Tr * (4328.31866498393 + Tr * (52325.3798350729 + -33255.7490812499 * Tr + 24515.5967903022 * Tr * Tr))) / (1 + 81.490354672861 * Tr)
    Case "R410a" 'Zerotropic HFC blend R-32+125 (50+.5,–1.5/50+1.5,–.5%), https://en.wikipedia.org/wiki/R-410A
        Tr = usrMaxDbl(0, 1 - TK / 344.51)
        Hc = 368.55
        Refrigerant_Hbubble = Hc - (Tr * (5201.29701026661 + Tr * (56110.3436425771 + -37951.9563805045 * Tr + 29802.5295009664 * Tr * Tr))) / (1 + 86.9160207163079 * Tr)
    Case "R600" 'CAS# 106-97-8: Butane, https://echa.europa.eu/substance-information/-/substanceinfo/100.003.136
        Tr = usrMaxDbl(0, 1 - TK / 425.13)
        Hc = 693.91
        Refrigerant_Hbubble = Hc - (Tr * (7474.46863484402 + Tr * (118031.599919433 + -73061.4013453035 * Tr + 36613.5643621172 * Tr * Tr))) / (1 + 82.244230350485 * Tr)
    Case "R600a" 'CAS# 75-28-5: Isobutane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.780
        Tr = usrMaxDbl(0, 1 - TK / 407.81)
        Hc = 633.94
        Refrigerant_Hbubble = Hc - (Tr * (5933.78729020458 + Tr * (81550.6034348199 + -46479.5335661856 * Tr + 19788.1064194294 * Tr * Tr))) / (1 + 62.6569962573858 * Tr)
    Case "R717" 'CAS# 7664-41-7: Ammonia (NH3), https://echa.europa.eu/substance-information/-/substanceinfo/100.028.760
        Tr = usrMaxDbl(0, 1 - TK / 405.4)
        Hc = 1119.22
        Refrigerant_Hbubble = Hc - (Tr * (19386.5537852179 + Tr * (141335.503781378 + -51644.2531815673 * Tr + 24630.8411918969 * Tr * Tr))) / (1 + 63.1166522083121 * Tr)
    Case "R744" 'CAS# 124-38-9: Carbon dioxide (CO2), https://echa.europa.eu/substance-information/-/substanceinfo/100.004.271
        Tr = usrMaxDbl(0, 1 - TK / 304.13)
        Hc = 332.25
        Refrigerant_Hbubble = Hc - (Tr * (8662.76845072689 + Tr * (122300.527720206 + -140235.701163124 * Tr + 169170.64073272 * Tr * Tr))) / (1 + 140.335963597642 * Tr)
    Case Else
        MsgBox "Refrigerant '" & Refrigerant & "' is not in database", vbExclamation
    End Select
End Function

Function Refrigerant_Hdew#(Refrigerant$, TK#)
    'Refrigerant_Hdew [kJ/kgK] = Dew-point enthalpy, i.e. right side of Ph-diagram
    'INPUTS:
    '- Refrigerant [string] = Refrigerant ASHRAE-code e.g. "R410a"
    '- TK [K] = Refrigerant temperature at dew-point [K]
    'SOURCE: Fitted to tables in ASHRAE Handbook Fundamentals, 2017, Chapter 30
    'NOTE: Equations generated by symbolic regression with https://github.com/SchildCode/RatFun-Regression:
    '  Average statistics for all 9 refrigerants: R² = 0.9988, Max Absolute Error = ±12.3 kJ/kg over the whole temperature range
    'AUTHOR: P.G.Schild, 2020

    Dim Tr# 'Complement to 1 of reducted temperature, i.e. 1 - (T [K] / Tcritical [K]), where Tcritical is the refrigerant's critical point
    Dim Hc# 'Enthalpy at refrigerant's critical point [kJ/(kgK)]

    Select Case Refrigerant
    Case "R32" 'CAS# 75-10-5: HFC "Difluoromethane", https://echa.europa.eu/substance-information/-/substanceinfo/100.000.764
        Tr = usrMaxDbl(0, 1 - TK / 351.26)
        Hc = 414.15
        Refrigerant_Hdew = Hc - (Tr * (-5932.29197274236 + Tr * (-7474.14778941814 + 37232.6789055002 * Tr - 22446.5110014522 * Tr * Tr))) / (1 + 54.8880924955789 * Tr)
    Case "R134a" 'CAS# 811-97-2: HFC "Norflurane" aka 1,1,1,2-Tetrafluoroethane, https://echa.europa.eu/substance-information/-/substanceinfo/100.011.252
        Tr = usrMaxDbl(0, 1 - TK / 374.21)
        Hc = 389.64
        Refrigerant_Hdew = Hc - (Tr * (-3751.70385782936 + Tr * (5831.65935922266 + 27700.362352493 * Tr - 24202.4102725855 * Tr * Tr))) / (1 + 66.8081879997468 * Tr)
    Case "R290" 'CAS# 74-98-6: HC Propane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.753
        Tr = usrMaxDbl(0, 1 - TK / 369.89)
        Hc = 555.24
        Refrigerant_Hdew = Hc - (Tr * (-6162.54497872838 + Tr * (14556.343263775 + 23814.1304213942 * Tr - 20162.9485399487 * Tr * Tr))) / (1 + 51.8069652067125 * Tr)
    Case "R407c" 'Zerotropic HFC blend R-32+125+134a (23±2/25±2/52±2%), https://en.wikipedia.org/wiki/R-407C
        Tr = usrMaxDbl(0, 1 - TK / 359.18)
        Hc = 378.48
        Refrigerant_Hdew = Hc - (Tr * (-4564.27373440949 + Tr * (882.872422159062 + 40394.4185405698 * Tr - 33600.2462786707 * Tr * Tr))) / (1 + 76.2370399341407 * Tr)
    Case "R410a" 'Zerotropic HFC blend R-32+125 (50+.5,–1.5/50+1.5,–.5%), https://en.wikipedia.org/wiki/R-410A
        Tr = usrMaxDbl(0, 1 - TK / 344.51)
        Hc = 368.55
        Refrigerant_Hdew = Hc - (Tr * (-3589.75407224144 + Tr * (-1361.38040952226 + 25709.0595297822 * Tr - 18483.9174340378 * Tr * Tr))) / (1 + 50.8134937179985 * Tr)
    Case "R600" 'CAS# 106-97-8: Butane, https://echa.europa.eu/substance-information/-/substanceinfo/100.003.136
        Tr = usrMaxDbl(0, 1 - TK / 425.13)
        Hc = 693.91
        Refrigerant_Hdew = Hc - (Tr * (-5333.19521865306 + Tr * (28454.2616111567 + 15346.8792540675 * Tr - 18640.3206761566 * Tr * Tr))) / (1 + 51.9950340400431 * Tr)
    Case "R600a" 'CAS# 75-28-5: Isobutane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.780
        Tr = usrMaxDbl(0, 1 - TK / 407.81)
        Hc = 633.94
        Refrigerant_Hdew = Hc - (Tr * (-5417.04632564831 + Tr * (27151.1755440779 + 18518.992936087 * Tr - 22659.4433334061 * Tr * Tr))) / (1 + 56.7071435554079 * Tr)
    Case "R717" 'CAS# 7664-41-7: Ammonia (NH3), https://echa.europa.eu/substance-information/-/substanceinfo/100.028.760
        Tr = usrMaxDbl(0, 1 - TK / 405.4)
        Hc = 1119.22
        Refrigerant_Hdew = Hc - (Tr * (-19517.5661343799 + Tr * (-46893.4973683591 + 164529.226824158 * Tr - 96822.8035355886 * Tr * Tr))) / (1 + 57.1634083848444 * Tr)
    Case "R744" 'CAS# 124-38-9: Carbon dioxide (CO2), https://echa.europa.eu/substance-information/-/substanceinfo/100.004.271
        Tr = usrMaxDbl(0, 1 - TK / 304.13)
        Hc = 332.25
        Refrigerant_Hdew = Hc - (Tr * (-9797.0760882652 + Tr * (-46687.6243854889 + 192043.685041829 * Tr - 201038.158033694 * Tr * Tr))) / (1 + 119.839331986652 * Tr)
    Case Else
        MsgBox "Refrigerant '" & Refrigerant & "' is not in database", vbExclamation
    End Select
End Function

Function Refrigerant_Tdew#(Refrigerant$, Pa#)
    'Refrigerant_Tdew [K] = Dew-point temperature at pressure Pa
    'INPUTS:
    '- Refrigerant [string] = Refrigerant ASHRAE-code e.g. "R410a"
    '- Pa [Pa] = Fluid pressure
    'SOURCE: Fitted to tables in ASHRAE Handbook Fundamentals, 2017, Chapter 30
    'NOTE: Equations generated by symbolic regression with https://github.com/SchildCode/RatFun-Regression:
    '  Average statistics for all 9 refrigerants: R² = 0.9999959, Max Absolute Error = ±1.02% over the whole temperature range
    'AUTHOR: P.G.Schild, 2020

    Dim LnP# 'Natural logarithm of pressure
    Dim Tr# 'Complement to 1 of reducted temperature, i.e. 1 - (T [K] / Tcritical [K]), where Tcritical is the refrigerant's critical point  => T = (1-Tr)*Tcritical

    LnP = Log(Pa)
    Select Case Refrigerant
    Case "R32" 'CAS# 75-10-5: HFC "Difluoromethane", https://echa.europa.eu/substance-information/-/substanceinfo/100.000.764
        Tr = (0.679844283196794 + LnP * (-5.58845879396756E-02 + 7.83846230683479E-04 * LnP)) / (1 + LnP * (-0.058449789491908 + 3.39983403637178E-05 * LnP * LnP))
        Refrigerant_Tdew = (1 - Tr) * 351.26
    Case "R134a" 'CAS# 811-97-2: HFC "Norflurane" aka 1,1,1,2-Tetrafluoroethane, https://echa.europa.eu/substance-information/-/substanceinfo/100.011.252
        Tr = (0.667225210249668 + LnP * (-6.06969167260545E-02 + 1.1065288032365E-03 * LnP)) / (1 + LnP * (-6.35333949333943E-02 + 4.54252966222158E-05 * LnP * LnP))
        Refrigerant_Tdew = (1 - Tr) * 374.21
    Case "R290" 'CAS# 74-98-6: HC Propane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.753
        Tr = (0.710757891310966 + LnP * (-6.85824911274988E-02 + 1.44282118828316E-03 * LnP)) / (1 + LnP * (-0.068752540282851 + 5.88631458267534E-05 * LnP * LnP))
        Refrigerant_Tdew = (1 - Tr) * 369.89
    Case "R407c" 'Zerotropic HFC blend R-32+125+134a (23±2/25±2/52±2%), https://en.wikipedia.org/wiki/R-407C
        Tr = (0.709718826194877 + LnP * (-8.28796136812859E-02 + 2.38656642321597E-03 * LnP)) / (1 + LnP * (-7.94321671852964E-02 + 8.16006897603657E-05 * LnP * LnP))
        Refrigerant_Tdew = (1 - Tr) * 359.18
    Case "R410a" 'Zerotropic HFC blend R-32+125 (50+.5,–1.5/50+1.5,–.5%), https://en.wikipedia.org/wiki/R-410A
        Tr = (0.706776530734329 + LnP * (-7.37091423758357E-02 + 1.80629447790992E-03 * LnP)) / (1 + LnP * (-7.14304225592101E-02 + 6.27474930697395E-05 * LnP * LnP))
        Refrigerant_Tdew = (1 - Tr) * 344.51
    Case "R600" 'CAS# 106-97-8: Butane, https://echa.europa.eu/substance-information/-/substanceinfo/100.003.136
        Tr = (0.687113813033299 + LnP * (-5.77922827938379E-02 + 8.19787892432219E-04 * LnP)) / (1 + LnP * (-6.01713491355452E-02 + 3.63711221709102E-05 * LnP * LnP))
        Refrigerant_Tdew = (1 - Tr) * 425.13
    Case "R600a" 'CAS# 75-28-5: Isobutane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.780
        Tr = (0.695367804879725 + LnP * (-6.18585825495433E-02 + 1.04667667420318E-03 * LnP)) / (1 + LnP * (-6.31658257395192E-02 + 4.3238673435457E-05 * LnP * LnP))
        Refrigerant_Tdew = (1 - Tr) * 407.81
    Case "R717" 'CAS# 7664-41-7: Ammonia (NH3), https://echa.europa.eu/substance-information/-/substanceinfo/100.028.760
        Tr = (0.720965429266934 + LnP * (-6.78109905326274E-02 + 1.44191817852209E-03 * LnP)) / (1 + LnP * (-6.53266151696419E-02 + 4.80245842200671E-05 * LnP * LnP))
        Refrigerant_Tdew = (1 - Tr) * 405.4
    Case "R744" 'CAS# 124-38-9: Carbon dioxide (CO2), https://echa.europa.eu/substance-information/-/substanceinfo/100.004.271
        Tr = (0.796087533695876 + LnP * (-9.46618938968321E-02 + 2.80263824593801E-03 * LnP)) / (1 + LnP * (-7.98925937377127E-02 + 7.70146828321376E-05 * LnP * LnP))
        Refrigerant_Tdew = (1 - Tr) * 304.13
    Case Else
        MsgBox "Refrigerant '" & Refrigerant & "' is not in database", vbExclamation
    End Select
End Function

Function Refrigerant_Pdew#(Refrigerant$, TK#)
    'Refrigerant_Pdew [Pa] = Stauration pressure given dew-point temperature
    'INPUTS:
    '- Refrigerant [string] = Refrigerant ASHRAE-code e.g. "R410a"
    '- TK [K] = Refrigerant temperature at dew-point [K]
    'SOURCE: Fitted to tables in ASHRAE Handbook Fundamentals, 2017, Chapter 30
    'NOTE: Equations generated by symbolic regression with https://github.com/SchildCode/RatFun-Regression:
    '  Average statistics for all 9 refrigerants: R² = 0.99998, Max Absolute Error = ±2.5% over the whole temperature range
    'AUTHOR: P.G.Schild, 2020

    Dim Tr# 'Complement to 1 of reducted temperature, i.e. 1 - (T [K] / Tcritical [K]), where Tcritical is the refrigerant's critical point

    Select Case Refrigerant
    Case "R32" 'CAS# 75-10-5: HFC "Difluoromethane", https://echa.europa.eu/substance-information/-/substanceinfo/100.000.764
        Tr = usrMaxDbl(0, 1 - TK / 351.26)
        Refrigerant_Pdew = Exp((15.5679015085583 + Tr * (-25.4660296027275 + 3.41848099187946 * Tr)) / (1 + Tr * (-1.17781759148254 + 0.199106729606217 * Tr * Tr)))
    Case "R134a" 'CAS# 811-97-2: HFC "Norflurane" aka 1,1,1,2-Tetrafluoroethane, https://echa.europa.eu/substance-information/-/substanceinfo/100.011.252
        Tr = usrMaxDbl(0, 1 - TK / 374.21)
        Refrigerant_Pdew = Exp((15.2139089053916 + Tr * (-25.5779302544142 + 3.41000371520309 * Tr)) / (1 + Tr * (-1.20224568507009 + 0.219248651387827 * Tr * Tr)))
    Case "R290" 'CAS# 74-98-6: HC Propane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.753
        Tr = usrMaxDbl(0, 1 - TK / 369.89)
        Refrigerant_Pdew = Exp((15.2724550053966 + Tr * (-25.7351896405739 + 5.65883114288467 * Tr)) / (1 + Tr * (-1.2400437455292 + 0.405468927677945 * Tr * Tr)))
    Case "R407c" 'Zerotropic HFC blend R-32+125+134a (23±2/25±2/52±2%), https://en.wikipedia.org/wiki/R-407C
        Tr = usrMaxDbl(0, 1 - TK / 359.18)
        Refrigerant_Pdew = Exp((15.3325241166998 + Tr * (-27.1862368342962 + 6.42493846609803 * Tr)) / (1 + Tr * (-1.26428115966836 + 0.542121403495998 * Tr * Tr)))
     Case "R410a" 'Zerotropic HFC blend R-32+125 (50+.5,–1.5/50+1.5,–.5%), https://en.wikipedia.org/wiki/R-410A
        Tr = usrMaxDbl(0, 1 - TK / 344.51)
        Refrigerant_Pdew = Exp((15.4031140650565 + Tr * (-25.4339253254685 + 3.46422969441002 * Tr)) / (1 + Tr * (-1.18439981904375 + 0.229629594514814 * Tr * Tr)))
   Case "R600" 'CAS# 106-97-8: Butane, https://echa.europa.eu/substance-information/-/substanceinfo/100.003.136
        Tr = usrMaxDbl(0, 1 - TK / 425.13)
        Refrigerant_Pdew = Exp((15.146247415958 + Tr * (-24.3915197064474 + 2.97147304394632 * Tr)) / (1 + Tr * (-1.16842463610541 + 0.158007448008371 * Tr * Tr)))
    Case "R600a" 'CAS# 75-28-5: Isobutane, https://echa.europa.eu/substance-information/-/substanceinfo/100.000.780
        Tr = usrMaxDbl(0, 1 - TK / 407.81)
        Refrigerant_Pdew = Exp((15.1025063119411 + Tr * (-24.204066495443 + 2.91816443423455 * Tr)) / (1 + Tr * (-1.16617077579373 + 0.165662127255591 * Tr * Tr)))
    Case "R717" 'CAS# 7664-41-7: Ammonia (NH3), https://echa.europa.eu/substance-information/-/substanceinfo/100.028.760
        Tr = usrMaxDbl(0, 1 - TK / 405.4)
        Refrigerant_Pdew = Exp((16.2418828732019 + Tr * (-25.7300624977046 + 3.02815984009489 * Tr)) / (1 + Tr * (-1.15415860983025 + 0.185327253381788 * Tr * Tr)))
    Case "R744" 'CAS# 124-38-9: Carbon dioxide (CO2), https://echa.europa.eu/substance-information/-/substanceinfo/100.004.271
        Tr = usrMaxDbl(0, 1 - TK / 304.13)
        Refrigerant_Pdew = Exp((15.8130484066556 + Tr * (-22.1723126027072 + 2.6653067678616 * Tr)) / (1 + Tr * (-0.971305513436541 + 0.54929879404179 * Tr * Tr)))
    Case Else
        MsgBox "Refrigerant '" & Refrigerant & "' is not in database", vbExclamation
    End Select
End Function

'----------------------------
'DRY AIR
'----------------------------

Public Function DryAir_Cp#(ByVal Tdry_K#)
    'DryAir_Cp [J/kgK] = Specific heat capacity of dry air, valid at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tdry_K [K] = Air dry-bulb temperature
    'NOTE: This is a fundamental property, based on empirical data

    If Tdry_K < 223.15 Or 473.15 < Tdry_K Then
        DryAir_Cp = ErrorMsg("Tdry_K out of range in DryAir_Cp.  Tdry_K = " & Tdry_K & " K")
    Else
        'REFERENCE: www.engineeringtoolbox.com
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.999160, Max deviation = ±0.379 J/kgK in the range 223K to 473K
        DryAir_Cp = (1 + Tdry_K ^ 2 * (-1.34590696095301E-04 - 6.28356110944288E-08 * Tdry_K)) / (Tdry_K * (1.18980667754983E-05 - 1.81101111121264E-07 * Tdry_K))
    End If
End Function

Public Function DryAir_Conduct#(ByVal Tdry_K#)
    'DryAir_Conduct [W/mK] = Specific heat capacity of dry air, valid at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tdry_K [K] = Air dry-bulb temperature
    'NOTE: This is a fundamental property, based on empirical data
    
    If Tdry_K < 225 Or 375 < Tdry_K Then
        DryAir_Conduct = ErrorMsg("Tdry_K out of range in DryAir_Conduct.  Tdry_K = " & Tdry_K & " K")
    Else
        'REFERENCE: www.engineeringtoolbox.com
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.999991, Max deviation = ±2.3E-05 W/mK in the range 225K to 375K
        DryAir_Conduct = Tdry_K * (9.72161872452033E-05 - 3.26055249878836E-08 * Tdry_K)
    End If
End Function

Public Function DryAir_DynaVisc#(ByVal Tdry_K#)
    'DryAir_DynaVisc [Pa·s] = Dynamic viscosity of dry air, valid at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tdry_K [K] = Air dry-bulb temperature
    'NOTE: This is a fundamental thermophysical property, based on empirical data, because it directly relates to the molecular characteristics and interactions in a fluid, independent of its density
    
    If Tdry_K < 223.15 Or 373.15 < Tdry_K Then
        DryAir_DynaVisc = ErrorMsg("Tdry_K out of range in DryAir_DynaVisc.  Tdry_K = " & Tdry_K & " K")
    Else
        'REFERENCE: www.engineeringtoolbox.com
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.9999925, Max deviation = ±1.27796E-08 Pa·s in the range 223K to 373K
        DryAir_DynaVisc = 7.95801050867814E-08 * Tdry_K / (1 + 9.7958282877158E-04 * Tdry_K)
    End If
End Function

Public Function DryAir_KineVisc#(ByVal Tdry_K#)
    'DryAir_KineVisc [m²/s] = Kinematic viscosity of dry air, valid at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tdry_K [K] = Air dry-bulb temperature
    'NOTE: This is strictly a derived property, governed by v = µ / rho, where v=kinematic viscosity [m²/s], µ=dynamic viscosity [Pa·s], rho=density [kg/m³]. However an explicit correlation is given here for efficiency, to avoid having to calculate v(T)=µ(T)/rho(T)

    If Tdry_K < 223.15 Or 373.15 < Tdry_K Then
        DryAir_KineVisc = ErrorMsg("Tdry_K out of range in DryAir_KineVisc.  Tdry_K = " & Tdry_K & " K")
    Else
        'REFERENCE: www.engineeringtoolbox.com
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.9999991, Max deviation = ±9.07384E-09 m²/s in the range 223K to 373K
        DryAir_KineVisc = 2.27048032645291E-10 * Tdry_K ^ 2 / (1 + 1.0085092689335E-03 * Tdry_K)
    End If
End Function

Public Function DryAir_Pr#(ByVal Tdry_K#)
    'DryAir_Pr [-] = Prandtl number of dry air, valid at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tdry_K [K] = Air dry-bulb temperature
    'NOTE: This is strictly a derived property, governed by:
    '  Pr =  µ cp / k
    'where:
    '  µ = absolute or dynamic viscosity [kg/(m s)]
    '  cp = specific heat [J/(kg K)]
    '  k = thermal conductivity [W/(m K)]
    'However, an explicit correlation is given here for efficiency, to avoid having to calculate Pr(T) =  µ(T)*cp(T)/k(T)

    If Tdry_K < 225 Or 375 < Tdry_K Then
        DryAir_Pr = ErrorMsg("Tdry_K out of range in DryAir_Pr.  Tdry_K = " & Tdry_K & " K")
    Else
        'REFERENCE: https://www.engineeringtoolbox.com/
        'Fitted by symbolic regression to a rational function using https://github.com/SchildCode/RatFun-Regression
        'Statistics: R²=0.99954, Max deviation = ±0.00045 in the range 225K to 375K
        DryAir_Pr = (1 + Tdry_K * (3.93117404779956E-02 - 4.19374808883644E-09 * Tdry_K ^ 2)) / 5.97982860437876E-02 * Tdry_K
    End If
End Function

Public Function DryAir_thermDiff#(ByVal Tdry_K#)
    'DryAir_thermDiff [m²/s] = Thermal diffusivity of dry air at standard atmospheric pressure (101325 Pa)
    'INPUTS: Tdry_K [K] = Air dry-bulb temperature
    'NOTE: This is a derived thermophysical property:
    '  a = k/(rho·Cp)  where a=thermal diffusivity [m²/s], k=thermal conductivity [W/(m·K)], rho=density [kg/m³], Cp=specific heat capacity [J/(kg·K)]

    If Tdry_K < 225 Or 375 < Tdry_K Then
        DryAir_thermDiff = ErrorMsg("Tw out of range in DryAir_thermDiff.  Tw = " & Tdry_K & " K")
    Else
        DryAir_thermDiff = DryAir_Conduct(Tdry_K) / (Air_DensH(Tdry_K, 0#, 101325#) * DryAir_Cp(Tdry_K))
    End If
End Function

'----------------------------
'MOIST AIR
'----------------------------

Public Function Air_Cp#(ByVal Tdry_K#, ByVal HumidRatio#)
    'Air_Cp [J/kgK] = Specific heat capacity of moist air
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- HumidRatio [kg/kg] = Humidity ratio [kg water vapour / kg dry air]

    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.9, Eqn.(27)
    Air_Cp = DryAir_Cp(Tdry_K) + HumidRatio * Vapour_Cp(Tdry_K) '[J/kgK]
End Function

Public Function Air_Enth#(ByVal Tdry_K#, ByVal HumidRatio#)
    'Air_Enth [J/kg dry air] = Air specific enthalpy per kg dry air
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- HumidRatio [kg/kg] = Humidity ratio [kg water vapour / kg dry air]
    'NOTE: This is strictly a derived thermophysical property, based on integrating dh = cp·dT over over the temperature range from reference state (T=0°C) to T_K. This is complicated by the fact that cp is itself a non-linear function of temperature. Therefore a dedicated correlation is given here

    Dim Tdry_C#
    
    Tdry_C = Tdry_K - ZeroC
    'REFERENCE: ASHRAE Handbook Fundamentals 2017, Page 1.2 and page 1.9 Eqn.(30)
    Air_Enth = (1.006 * Tdry_C + HumidRatio * (2501 + 1.86 * Tdry_C)) * 1000 '[J/kg]
End Function

Public Function Air_DryBulb#(ByVal Enthalpy_Jkg#, ByVal HumidRatio#)
    'Air_DryBulb [K] = Dry-bulb air temperature
    'INPUTS:
    '- Enthalpy_Jkg [J/kg] = Air total enthalpy
    '- HumidRatio [kg/kg]= Humidity ratio [kg water vapour / kg dry air]
    
    Dim Tdry_C#
    
    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.2 and page 1.9 Eqn.(30) inverted
    Tdry_C = (Enthalpy_Jkg - HumidRatio * 2501000#) / (1006# + HumidRatio * 1860)
    Air_DryBulb = Tdry_C + ZeroC '[degrees Kelvin]
End Function

Public Function Air_DensH#(ByVal Tdry_K#, ByVal HumidRatio#, ByVal Patm_Pa#)
    'Air_DensH [kg/m³] = Air density, given humidity ratio
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- HumidRatio [kg/kg] = Humidity ratio [kg water vapour / kg dry air]
    '- Patm_Pa [Pa] = Atmospheric pressure
    
    'REFERENCE: ASHRAE Handbook Fundamentals 2017, eqns (9b) and (26)
    'Possible alternative: R.S. Davis, "Equation for the Determination of the Density of Moist Air (1981/91)", Metrologia, 1992, iop.org
    Air_DensH = Patm_Pa / (287.042 * Tdry_K * (1 + 1.607858 * HumidRatio)) * (1 + HumidRatio) '[kg/m3]
End Function

Public Function Air_DensR#(ByVal Tdry_K#, ByVal RH#, ByVal Patm_Pa#)
    'Air_DensR [kg/m³] = Air density, given RH
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- RH [-] = Relative humidity [0 <= fraction <= 1]
    '- Patm_Pa [Pa] = Atmospheric pressure
    
    Dim HumidRatio#
    
    'REFERENCE: ASHRAE Handbook Fundamentals 2017
    HumidRatio = Air_HumidRatioR(Tdry_K, RH, Patm_Pa)  '[kg/kg]
    Air_DensR = Air_DensH(Tdry_K, HumidRatio, Patm_Pa) '[kg/m3]
End Function

Public Function Air_TdewP#(ByVal Tdry_K#, ByVal Pv_Pa#)
    'Air_TdewP [K] = Dew-point (or frost-point) temperature, given parial pressure of vapour
    'NOTE: This function is basically an inverse of function Vapour_Pws(), using Newton Raphson iteration
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature. Used here only as first guess of Tdew
    '- Pv_Pa [Pa] = Water vapour partial pressure

    Dim LnP#
    Dim iter&
    Dim T0#
    Dim T1#
    Dim f0#
    Dim f1# 'deviation Vapour_Pws(Td1)-Pv_Pa [Pa]
    
    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.10 Eqn.(37) & (38) and refined by root-finding function Vapour_Pws() with Secant method
    'First approximation of Tdew
    LnP = Log(Pv_Pa * 0.001)
    If Tdry_K < ZeroC Then
        T1 = ZeroC + 6.09 + LnP * (12.608 + LnP * 0.4959) 'Eqn.(38) below 0°C
    Else
        T1 = ZeroC + 6.54 + LnP * (14.526 + LnP * (0.7389 + LnP * 0.09486)) + 0.4569 * (Pv_Pa * 0.001) ^ 0.1984 'Eqn.(37) between 0°C and 93°C
    End If
    'Now refine using Secant method
    'Benefits of Secant method: (i) almost quadratic convergence, (ii) no need to calculate derivatives, and (iii) converges from a wider range of starting points than Newton't method
    T0 = Tdry_K  'upper limit of Tdew
    f0 = Vapour_Pws(T0) - Pv_Pa
    f1 = Vapour_Pws(T1) - Pv_Pa
    For iter = 1 To 100 'Typically only 4 iterations needed
        Air_TdewP = T1 - f1 * (T1 - T0) / (f1 - f0) 'Secant correction, in a form that minimizes roundoff error
        If Tdry_K < Air_TdewP Then Air_TdewP = Tdry_K
        If Abs(T1 - T0) < 0.001 Then Exit For
        T0 = T1
        T1 = Air_TdewP
        f0 = f1
        f1 = Vapour_Pws(Air_TdewP) - Pv_Pa
    Next
End Function

Public Function Air_TdewH#(ByVal Tdry_K#, ByVal HumidRatio#, ByVal Patm_Pa#)
    'Air_TdewH [K] = Dew-point (or frost-point) temperature, given humidity ratio
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- HumidRatio [kg/kg] = Humidity ratio [kg water vapour / kg dry air]
    '- Patm_Pa [Pa] = Atmospheric pressure

    Dim Pv# 'Vapour pressure (saturation pressure at dew-point)

    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.10 Eqn.(36)
    Pv = Patm_Pa * HumidRatio / (0.621945 + HumidRatio)
    Air_TdewH = Air_TdewP(Tdry_K, Pv)
End Function

Public Function Air_TdewR#(ByVal Tdry_K#, ByVal RH#)
    'Air_TdewR [K] = Dew-point (or frost-point) temperature, given RH
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- RH [-] = Relative humidity = Vapour_Pws(Tdew) / Vapour_Pws(Ta)  =>  Vapour_Pws(Tdew) = Vapour_Pws(Ta) * RH  =>  Tdew = Air_TdewP[Vapour_Pws(Tdew)]

    Dim Pv# '[Pa] Vapour pressure

    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.8 Eqn.(12)
    Pv = Vapour_Pws(Tdry_K) * RH
    Air_TdewR = Air_TdewP(Tdry_K, Pv)
End Function

Public Function Air_TdewW#(ByVal Tdry_K#, ByVal Twet_K#, ByVal Patm_Pa#)
    'Air_TdewW [K] = Dew-point (or frost-point) temperature, given wet-bulb temperature
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- Twet_K [K] = Wet-bulb temperature
    '- Patm_Pa [Pa] = Atmospheric air pressure

    Dim HumidRatio# '[-] = W = [kg water vapour / kg dry air]

    'REFERENCE: ASHRAE Handbook Fundamentals 2017
    HumidRatio = Air_HumidRatioW(Tdry_K, Twet_K, Patm_Pa)
    Air_TdewW = Air_TdewH(Tdry_K, HumidRatio, Patm_Pa)
End Function

Public Function Air_HumidRatioR#(ByVal Tdry_K#, ByVal RH#, ByVal Patm_Pa#)
    'Air_HumidRatioR [kg water vapour / kg dry air] = Humidity ratio of moist air, given RH
    'INPUTS:
    '- Tdry_K [K] = dry-bulb air temperature
    '- RH [-] = Relative humidity (over liquid water)
    '- Patm_Pa [Pa] = Atmospheric air pressure

    Dim Pv# '[Pa] Vapour pressure

    If RH <= 0 Then
        Air_HumidRatioR = 0#
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.9 Eqn.(20)
        Pv = RH * Vapour_Pws(Tdry_K)
        Air_HumidRatioR = 0.621945 * Pv / (Patm_Pa - Pv)
    End If
End Function

Public Function Air_HumidRatioD#(ByVal Tdew_K#, ByVal Patm_Pa#)
    'Air_HumidRatioD [kg water vapour / kg dry air] = Humidity ratio of moist air, given dew-point temperature
    'INPUTS:
    '- Tdew_K [K] = Air dew-point (or frost-point if over ice below 0°C) temperature
    '- Patm_Pa [Pa] = Atmospheric air pressure

    Dim Pvs# 'Saturated vapour pressure at dew-point [Pa]
  
    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.9 Eqn.(20)
    Pvs = Vapour_Pws(Tdew_K, False) '[Pa] Option TRUE if Tdew_K is measured with a chilled mirror hygrometer, as ice can form below 0°C. Default is FALSE
    Air_HumidRatioD = 0.621945 * Pvs / (Patm_Pa - Pvs)
End Function

Public Function Air_HumidRatioW#(ByVal Tdry_K#, ByVal Twet_K#, ByVal Patm_Pa#)
    'Air_HumidRatioW [kg water vapour / kg dry air] = Humidity ratio of moist air, given wet-bulb tempeature
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- Twet_K [K] = Air wet-bulb temperature
    '- Patm_Pa [Pa] = Atmospheric air pressure

    Dim HumidRatio#
    Dim Wsat# 'Humidity ratio at saturation, at given web bulb temp.
    Dim Pvs#
    Dim Tdry_C#
    Dim Twet_C#

    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.9 Eqn.(33) & (35)
    Tdry_C = Tdry_K - ZeroC
    Twet_C = Twet_K - ZeroC
    Pvs = Vapour_Pws(Twet_K) '[Pa] Saturation vapour pressure at wet-bulb temperature
    Wsat = 0.621945 * Pvs / (Patm_Pa - Pvs) '[kg/kg] eqn.(21)
    If Twet_K < 273.16 Then
        'Alternative formulation for ice-bulb temperature
        Air_HumidRatioW = ((2830 - 0.24 * Twet_C) * Wsat - 1.006 * (Tdry_C - Twet_C)) / (2830 + 1.86 * Tdry_C - 2.1 * Twet_C) 'eqn.(35)
    Else
        'Note: equation is strictly for water, as is the convention for meteo calculations.
        Air_HumidRatioW = ((2501 - 2.326 * Twet_C) * Wsat - 1.006 * (Tdry_C - Twet_C)) / (2501 + 1.86 * Tdry_C - 4.186 * Twet_C) 'eqn.(33)
    End If
End Function

Public Function Air_HumidRatioX#(ByVal SpecificHumid#)
    'Air_HumidRatioX [kg water vapour / kg dry air] = Humidity ratio of moist air, given specific humidity
    'INPUTS: SpecificHumidity [kg/kg] = X = [kg water vapour / kg moist air]

    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.8 Eqn.(9b)
    Air_HumidRatioX = SpecificHumid / (1# - SpecificHumid)
End Function

Public Function Air_RHd#(ByVal Tdry_K#, ByVal Tdew_K#, Optional ice As Boolean = False)
    'Air_RHd [-] = Relative humidity (over liquid water), given dew-point temperature
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- Tdew_K [K] = Air dew-point temperature (or frost-point when below 0°C, if option "ice"=TRUE)
    '- ice = TRUE if over ice when below 0°C (e.g. chilled mirror hygrometer), default is FALSE (Meteo data is always over liquid water)
    'NOTE: This routine applies the fundamental definition of relative humidity.

    Dim Pv# 'Vapour pressure (partial pressure at Tdry) = Saturation vapour pressure at dew-point (or frost-point)
    Dim Pvs# 'Saturation vapour pressure
    
    'Note on Vapour_Pws#(Tdew_K, TRUE or FALSE):
    ' FALSE (default): In meteorology, RH is always defined as humidity over a plane of liquid water (not ice).
    ' TRUE if Tdew is measured with a chilled mirror hygrometer, as ice can form below 0°C.
    'REFERENCE: Richardson, Knuteson & Tobin, "A Chilled Mirror Dew-point Hygrometer for Field Use", 9th ARM Science Team Meeting Proceedings, San Antonio, Texas, March 22-26, 1999
    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.8 Eqn.(12)
    Pv = Vapour_Pws(Tdew_K, False) 'Option TRUE if Tdew is measured with a chilled mirror hygrometer, as ice can form below 0°C. Default is FALSE
    Pvs = Vapour_Pws(Tdry_K)
    Air_RHd = Pv / Pvs
End Function

Public Function Air_RHh#(ByVal Tdry_K#, ByVal HumidRatio#, ByVal Patm_Pa#)
    'Air_RHh [-] = Relative humidity (over liquid water), given humidity ratio
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- Humidratio [kg/kg] = [kg water vapour / kg dry air]
    '- Patm_Pa [Pa] = Atmospheric air pressure

    Dim Pv# 'Vapour pressure (partial pressure at Tdry) = Saturation vapour pressure at dew-point (or frost-point)
    Dim Pvs# 'Vapour pressure at saturation
    
    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.8 Eqn.(12) and (20)
    Pv = Patm_Pa * HumidRatio / (0.621945 + HumidRatio)
    Pvs = Vapour_Pws(Tdry_K)
    Air_RHh = Pv / Pvs
End Function

Public Function Air_RHw#(ByVal Tdry_K#, ByVal Twet_K#, ByVal Patm_Pa#)
    'Air_RHw [-] = Relative humidity (over liquid water), given wet-bulb temperature
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- Twet_K [K] = Air wet-bulb temperature
    '- Patm_Pa [Pa] = Atmospheric air pressure

    Dim HumidRatio# 'Humidity ratio [kg water vapour / kg dry air]
    
    'REFERENCE: ASHRAE Handbook Fundamentals 2017
    HumidRatio = Air_HumidRatioW(Tdry_K, Twet_K, Patm_Pa)
    Air_RHw = Air_RHh(Tdry_K, HumidRatio, Patm_Pa)
End Function

Public Function Air_TwetH#(ByVal Tdry_K#, ByVal HumidRatio#, ByVal Patm_Pa#)
    'Air_TwetH [K] = Wet-bulb temperature, given humidity ratio. Fast iterative reverse calculation by secant method
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- HumidRatio [kg/kg] = [kg water vapour / kg dry air]
    '- Patm_Pa [Pa] = Atmospheric air pressure

    Dim Tdew_K#
    Dim iter&
    Dim T0# 'Twet guess [K]
    Dim T1# 'Twet guess [K]
    Dim f0# 'Error function Air_HumidRatioW(T0) - HumidRatio
    Dim f1# 'Error function Air_HumidRatioW(T1) - HumidRatio
    Dim f2# 'Error function Air_HumidRatioW(Air_TwetH) - HumidRatio

    Tdew_K = Air_TdewH(Tdry_K, HumidRatio, Patm_Pa)
    If Tdry_K < Tdew_K Then
        Air_TwetH = ErrorMsg("Tdry_K is lower than Tdew_K in Function Air_TwetH")
    Else
        'REFERENCE: ASHRAE Handbook Fundamentals 2017
        'We can safely use the Secant method on well-behaved function Vapour_Pws
        'Benefits of Secant method: (i) almost quadratic convergence, (ii) no need to calculate derivatives, and (iii) converges from a wider range of starting points than Newton't method
        T0 = Tdew_K 'lower bound
        T1 = 0.5 * (Tdew_K + Tdry_K) '2nd guess
        f0 = Air_HumidRatioW(Tdry_K, T0, Patm_Pa) - HumidRatio
        f1 = Air_HumidRatioW(Tdry_K, T1, Patm_Pa) - HumidRatio
        For iter = 1 To 100 'Typically only 4 iterations needed, a limit of 100 absolutely guarantees convergence
            Air_TwetH = T1 - f1 * (T1 - T0) / (f1 - f0) 'Secant correction, in a form that minimizes roundoff error
            If Abs(T1 - T0) < 0.001 Then Exit For
            T0 = T1
            T1 = Air_TwetH
            f0 = f1
            f1 = Air_HumidRatioW(Tdry_K, Air_TwetH, Patm_Pa) - HumidRatio
        Next
    End If
End Function

Public Function Air_TwetR#(ByVal Tdry_K#, ByVal RH#, ByVal Patm_Pa#)
    'Air_TwetR [K] = Wet-bulb temperature, given RH
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- RH [-] = Relative humidity
    '- Patm_Pa [Pa] = Atmospheric air pressure

    Dim HumidRatio# 'HumidRatio = W = [kg water vapour / kg dry air]

    'REFERENCE: ASHRAE Handbook Fundamentals 2017
    HumidRatio = Air_HumidRatioR(Tdry_K, RH, Patm_Pa)
    Air_TwetR = Air_TwetH(Tdry_K, HumidRatio, Patm_Pa)
End Function

Public Function Air_TwetD#(ByVal Tdry_K#, ByVal Tdew_K#, ByVal Patm_Pa#)
    'Air_TwetD [K] = Wet-bulb temperature, given dew-point temperature
    'INPUTS:
    '- Tdry_K [K] = Air dry-bulb temperature
    '- Tdew_K [K] = Air dew-point temperature
    '- Patm_Pa [Pa] = Atmospheric air pressure

    Dim HumidRatio# 'HumidRatio = W = [kg water vapour / kg dry air]

    'REFERENCE: ASHRAE Handbook Fundamentals 2017
    HumidRatio = Air_HumidRatioD(Tdew_K, Patm_Pa)
    Air_TwetD = Air_TwetH(Tdry_K, HumidRatio, Patm_Pa)
End Function

'------------------------------
'CLIMATE/ATMOSPHERIC PROPERTIES
'------------------------------

Public Function Atmos_Pa#(ByVal Height_over_sea_level_m#)
    'Atmos_Pa [Pa] = Estimate of standard-atmosphere barometric pressure as a function of height over sea level
    'INPUTS: Height_over_sea_level_m [m], e.g. height of building or meteo station

    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.1, Eqn.(3)
    Atmos_Pa = 101325 * (1 - 0.0000225577 * Height_over_sea_level_m) ^ 5.2559 'Curve fitted to NASA tabulated values at 0m and 11km
End Function

Public Function Atmos_Pa2#(ByVal Atmos_Pa1#, ByVal Atmos_T1#, ByVal Altitude1_m#, ByVal Altitude2_m#)
    'Atmos_Pa2 [Pa] = Estimate of atmospheric pressure at altitude 1 (e.g. sea level, og building site) given pressure and temperature at Altitude2 (e.g. station altitude)
    'NOTE: It takes into account vertical temperature profile, assuming typical moisture content (unsaturated environmental lapse rate).
    'INPUTS:
    '- Atmos_Pa1 [Pa] = Measured barometric pressure at Altitude1
    '- Atmos_T1 [K] = Measured dry-bulb temperature at Altitude1
    '- Altitude1_m [m] = Altitude at which Atmos_Pa1 and Atmos_T1 are measured
    '- Altitude2_m [m] = Altitude for which you wish to estimate the barometric pressure. Can be higher or lower than Altitude1
    'REFERENCE: https://en.wikipedia.org/wiki/Barometric_formula

    Const ELR# = -0.0065 'Environmental lapse rate, 0.65°C per 100 m in troposhere according to International Standard Atmosphere (https://en.wikipedia.org/wiki/International_Standard_Atmosphere)
    Const molM# = 0.0289644 'molar mass of Earth's air: 0.0289644 kg/mol
    Const uniR# = 8.3144598 'Universal gas constant: 8.3144598  J /mol/K
    Dim Tdry2_K#

    Tdry2_K = Atmos_T2(Atmos_T1, Altitude1_m, ByVal Altitude2_m)
    'Estimates barometric pressure as a function of height over sea level. valid in troposhere, up to 11 km.
    Atmos_Pa2 = Atmos_Pa1 * (Atmos_T1 / Tdry2_K) ^ (9.80665 * molM / (uniR * ELR)) 'barometric formula
End Function

Public Function Atmos_T#(ByVal Height_over_sea_level_m#)
    'Atmos_T [K] = Estimate temperature as a function of height over sea level. Valid in troposphere (<11km)
    'INPUTS: Height_over_sea_level_m [m], e.g. height of building or meteo station
    'NOTE: This assumes typical moisture content (unsaturated adiabatic lapse rate)
    '  Dry adiabatic lapse rate (DALR) is steeper = g/cp = 9.81/1.006 = 0.009748 °C/m, or about 1 °C/100m
    'REFERENCE: ASHRAE Handbook Fundamentals 2017, page 1.1 Eqn.(4)
    
    Const ELR# = -0.0065 'Environmental lapse rate, 0.65°C per 100 m in troposhere according to International Standard Atmosphere (https://en.wikipedia.org/wiki/International_Standard_Atmosphere)
    
    Atmos_T = 288.14 + ELR * Height_over_sea_level_m
End Function

Public Function Atmos_T2#(ByVal Atmos_T1, ByVal Altitude1_m#, ByVal Altitude2_m#)
    'Atmos_T2 [K] = Estimate of dry-bulb temperature at Altitude2_m given Tdry1_K at Altitude1_m. Assumes Environmental Lapse Rate (ELR)
    'INPUTS:
    '- Atmos_T1 [K] = Measured dry-bulb temperature at Altitude1
    '- Altitude1_m [m] = Altitude at which Atmos_T1 is measured
    '- Altitude2_m [m] = Altitude for which you wish to estimate temperature Atmos_T2 . Can be higher or lower than Altitude1
    'REFERENCE: ASbHRAE Handbook Fundamentals 2017, page 1.1 Eqn.(4)

    Const ELR# = -0.0065 'Environmental lapse rate, 0.65°C per 100 m in troposhere according to International Standard Atmosphere (https://en.wikipedia.org/wiki/International_Standard_Atmosphere)

    Atmos_T2 = Atmos_T1 + ELR * (Altitude2_m - Altitude1_m)
End Function

Public Function Wind_Loc2_ms#(ByVal Wind_Loc1_ms#, ByVal Alpha_Loc1#, ByVal Alpha_Loc2#, ByVal Height_Loc2_m#)
    'Wind_Loc2_ms [m/s] = Wind speed at Location 2 given wind speed from 10 m high weather station at Location1.
    'INPUTS:
    '- Wind_Loc1_ms [m/s] = Wind velocity at location 1 (weather station) at height 10m
    '- Alpha_Loc1 [-] = Power-law wind profile exponent at location 1 (weather station)
    '- Alpha_Loc2 [-] = Power-law wind profile exponent at location 1 (local site)
    '- Height_Loc2_m [m = Height above ground of wind speed at location 2
    'NOTE: Assumes power-law wind profile. Assumes that all profiles have same gradient wind velocity above boundary layer, over approx 400 m
    '  Not quite as accurate as log-law wind velocity profile, but good for practical use
    'REFERENCE:
    '- http://en.wikipedia.org/wiki/Wind_profile_power_law
    
    'Alpha =
    '0.100  Ocean or other body of water with at least 5km of unrestricted expanse
    '0.118  Flat terrain - no obstacles, beach, ice plain, snow field
    '0.149  Open terrain - low grass, field without crop (fallow land)
    '0.150  Flat terrain with some isolated ostacles, e.g. buildings/trees well separated from each other
    '0.160  Flat open country
    '0.170  Open flat country
    '0.182  Roughly open - low crops, low hedges, few trees, very few houses
    '0.200  Rural/countrside with scattered wind breaks (e.g. low buildings, trees)
    '0.218  Rough - high and low crops, large obstacles at distances 15*H, rows of trees, low orchards
    '0.220  Rolling or level surface broken by numerous obstructions, such as trees or small houses
    '0.250  Urban, industrial or forrest landscape
    '0.257  Very rough - obstacles at distances of 10*H, spread wood, farm buildings, vineyards
    '0.313  Closed landscape
    '0.330  City landscape
    '0.350  Centre of large city
    '0.377  City centre - alternated low and highrise
    
    Dim Hg1# 'Gradient height at met station [m]
    Dim Hg2# 'Gradient height at building site [m]

    '(1) Estimate gradient heights. Correlation based on Davenport data from 1960, also used in Standard ASCE 7
    Hg1 = -2012.1 * Alpha_Loc1 ^ 2 + 1919.4 * Alpha_Loc1 + 42.444
    Hg2 = -2012.1 * Alpha_Loc2 ^ 2 + 1919.4 * Alpha_Loc2 + 42.444
    '(2) Calculate local velocity at given building height
    Wind_Loc2_ms = Wind_Loc1_ms * (Hg1 / 10) ^ Alpha_Loc1 * (Height_Loc2_m / Hg2) ^ Alpha_Loc2
End Function

'----------------------------
'FLUID FLOW
'----------------------------

Public Function OrificeMassFlow_m3s(Tappings_str$, DuctDia_m#, OrificeDia_m#, Tdry_K#, RH#, Patm_Pa#, dP_Pa#) As Variant
    'OrificeMassFlow_m3s [m³/s] =  Volumetric air flow rate measured by means of pressure drop over an orifice plate
    'INPUTS:
    '- Tappings_str = "Corner" for corner tappings, "D & D/2" for D & D/2 tappings, "Flange" for flange tappings
    '- DuctDia_m [m] = Upstream internal duct diameter at working conditions (D)
    '- OrificeDia_m [m] = Diameter of ISO 5167-1 standard orifice at working conditions (d)
    '- Tdry_K [K] = Dry-bulb air temperature upstream of ISO orifice plate
    '- RH [-] = Relative humidity upstream of ISO orifice plate [0 to 1]
    '- Patm_Pa [Pa] = Atmospheric air pressure upstream of ISO orifice plate
    '- dP_Pa [Pa] = Differential pressure measured over ISO standard orifice
    'REFERENCE: ISO/FDIS 5167-2:2002
    'Author: Peter.Schild@oslomet.no, 2020

    Const Kappa# = 1.4 'Isentropic exponent of air [-]
    Dim e# 'Expansibility factor upstream [-]
    Dim Beta# 'Diameter ratio [-]
    Dim mu# 'Dynamic viscosity upstream of ISO orifice plate [Pa·s]
    Dim rho# 'Air density upstream of ISO orifice plate [kg/m³]
    Dim HumidRatio# 'Humidity ratio upstream of ISO orifice plate [kg/kg]
    Dim C# 'Discharge coefficient for ISO orifice plate, calculated using Reader-Harris/Gallagher equation, depends on Re [-]
    Dim C_infinite# 'Discharge coefficient for ISO orifice plate, when Re is infinite [-]
    Dim L1# 'Quotient of distance of upstream tapping from upstream orifice face, and pipe diameter [m]
    Dim L2# 'Quotient of distance of downstream tapping from downstream orifice face, and pipe diameter [m]
    Dim dummy# 'Dummy variable
    Dim Re# 'Reynolds no. at flow diameter DuctDia_m [nondimensional]
    Dim Invariant# 'Used to speed up iteration loop
    Dim a# 'Parameter in Reader-Harris/Gallaghet equation (C)
    Dim M2# 'Parameter in Reader-Harris/Gallaghet equation (C)
    Dim ErrTxt$ 'Error message
        
    'Pre-checks for limits of applicability
    ErrTxt = ""
    If OrificeDia_m = 0 Then ErrTxt = "Dia?" 'This row is empty. Just skip
    If DuctDia_m < 0.05 Then ErrTxt = "D<5cm!"
    If 1 < DuctDia_m Then ErrTxt = "D>1m!"
    If OrificeDia_m < 0.0125 Then ErrTxt = "d<12.5cm!"
    If ErrTxt <> "" Then GoTo jumpErr
    Beta = OrificeDia_m / DuctDia_m
    If Beta < 0.1 Then ErrTxt = "d/D< 0.1!"
    If 0.75 < Beta Then ErrTxt = "d/D>0.75!"
    If (Patm_Pa - dP_Pa) / Patm_Pa < 0.75 Then ErrTxt = "high dP!"
    If ErrTxt <> "" Then GoTo jumpErr

    'Precalculations
    e = 1# - (0.351 + 0.256 * Beta ^ 4# + 0.93 * Beta ^ 8#) * (1# - ((Patm_Pa - dP_Pa) / Patm_Pa) ^ (1# / 1.4))
    HumidRatio = Air_HumidRatioR(Tdry_K, RH, Patm_Pa)
    rho = Air_DensH(Tdry_K, HumidRatio, Patm_Pa)
    mu = DryAir_DynaVisc(Tdry_K)
    Invariant = e * OrificeDia_m ^ 2 * Sqr(2 * dP_Pa * rho) / (mu * DuctDia_m * Sqr(1 - Beta ^ 4))
    Select Case Tappings_str
        Case "Corner" 'Corner tappings
            L1 = 0
            L2 = 0
        Case "D & D/2" 'D & D/2' tappings
            L1 = 1
            L2 = 0.47
        Case "Flange" 'Flange tappings
            L1 = 0.0254 / DuctDia_m
            L2 = L1
        Case Else 'error
            ErrTxt = "Tapping type?"
            GoTo jumpErr
    End Select
    M2 = 2# * L2 / (1# - Beta)
    dummy = (0.043 + 0.08 * Exp(-10# * L1) - 0.123 * Exp(-7# * L1)) * Beta ^ 4 / (1# - Beta ^ 4#)
    C_infinite = 0.5961 + 0.0261 * Beta ^ 2# - 0.216 * Beta ^ 8# + dummy - 0.031 * (M2 - 0.8 * M2 ^ 1.1) * Beta ^ 1.3
    If DuctDia_m < 0.07112 Then C_infinite = C_infinite + 0.011 * (0.75 - Beta) * (2.8 - DuctDia_m / 0.0254)
    'First guess of C, assuming Re is infinite
    C = C_infinite
    
    'Iteration
    Do
        Re = C * Invariant 'Reynolds number of flow at diameter DuctDia_m
        a = (19000 * Beta / Re) ^ 0.8
        'Next guess of C, assuming Re
        C = C_infinite + 0.000521 * (1000000# * Beta / Re) ^ 0.7 + (0.0188 + 0.0063 * a) * Beta ^ 3.5 * (1000000# / Re) ^ 0.3 + -0.11 * a * dummy
    Loop Until Abs((Invariant - Re / C) / Invariant) < 0.00000001

   'Post-checks for limits of applicability for Re
    Select Case Tappings_str
        Case "Corner", "D & D/2" 'Corner' or 'D & D/2' tappings
            If Beta < 0.56 Then
                If Re < 5000 Then ErrTxt = "Re<5000!"
            Else
                dummy = 16000 * Beta ^ 2
                If Re < dummy Then ErrTxt = "Re<" & CStr(Int(dummy)) & "!"
            End If
        Case Else 'Flange' tappings
            If Re < 5000 Then
                ErrTxt = "Re<5000!"
            Else
                dummy = 170 * Beta ^ 2 * (DuctDia_m * 1000#)
                If Re < dummy Then ErrTxt = "Re<" & CStr(Int(dummy)) & "!"
            End If
    End Select
jumpErr:
    If ErrTxt = "" Then
        OrificeMassFlow_m3s = CDbl(onePi * mu * DuctDia_m * Re / (4 * rho)) 'volumetric units
        'OrificeMassFlow_kgps = CDbl(onePi * mu * DuctDia_m * Re / 4) 'gravimetric units
    Else
        OrificeMassFlow_m3s = CVar(ErrTxt)
    End If
End Function

Public Function Air_DuctFriction#(FlowRate_m3s#, Diam_m#, Roughness_m#, Tdry_K#, RH#, Patm_Pa#)
    'Air_DuctFriction [Pa/m] = Pressure drop per meter duct, for airflow in ducts
    'INPUTS:
    '- FlowRate_m3s [m³/s] = Air volume flow
    '- Diam_m [m] = Internal diameter of round duct (or hydraulic diameter of non-round duct)
    '- Roughness_m [m] = RMS surface roughness of duct
    '- Tdry_K [K] = Dry-bulb air temperature of bulk air flow
    '- RH [-] = Relative humidity upstream of bulk air flow
    '- Patm_Pa [Pa] = Static pressure in duct
    'Note: To get standard a density of 1.2 kg/m³, use Tdry=293.15, Patm=101325, and RH=0.395122
    'Note: Typical spiro-duct roughness: 3.53E-7 m for ducts <= 200 mm diameter, and 1.77E-4 for ducts > 200 mm
    'Author: Peter.Schild@oslomet.no, 2020

    Dim ff# 'Darcy friction factor [-]
    Dim Pdyn# 'Dynamic pressure [Pa]
    Dim Re# 'Reynolds number [-]
    Dim darea# 'circulat pipe or duct cross sectionsl flow area [m²]
    Dim rho# 'Density of air [kg/m³]
    Dim HumidRatio# 'Humidity ratio [-]
    Dim vv# 'Kinematic viscosity [m²/s]
    Dim uu# 'Velocity [m/s]
    
    'Air properties
    HumidRatio = Air_HumidRatioR(Tdry_K, RH, Patm_Pa)
    rho = Air_DensH(Tdry_K, HumidRatio, Patm_Pa)
    vv = DryAir_KineVisc(Tdry_K)

    'Flow properties
    darea = 0.25 * onePi * Diam_m * Diam_m 'Cross section area [m²]
    uu = FlowRate_m3s / darea 'nominal velocity [m/s]
    Re = uu * Diam_m / vv 'Reynolds number [-]
    ff = FrictionFactor(Re, Roughness_m / Diam_m) 'Darcy friction factor [-]
    Pdyn = 0.5 * rho * uu * uu 'Dynamic pressure [Pa]
    Air_DuctFriction = ff * Pdyn / Diam_m '[Pa/m]. Darcy-Weisbach equation: dP/L = f/D*Pdyn
End Function

Public Function Water_PipeFriction#(FlowRate_m3s#, Diam_m#, Roughness_m#, T_K#)
    'Water_PipeFriction_Pa [Pa/m] = Pressure drop per meter pipe, for water flow in pipes
    'INPUTS:
    '- FlowRate_m3s [m³/s] = Water volume flow
    '- Diam_m [m] = Internal diameter of round pipe (or hydraulic diameter of non-round pipe)
    '- Roughness_m [m] = RMS surface roughness of pipe
    '- T_K [K] = Temperature of bulk water flow
    'Author: Peter.Schild@oslomet.no, 2020

    Dim ff# 'Darcy friction factor [-]
    Dim Pdyn# 'Dynamic pressure [Pa]
    Dim Re# 'Reynolds number [-]
    Dim darea# 'circulat pipe or pipe cross sectionsl flow area [m²]
    Dim rho# 'Density of water [kg/m³]
    Dim HumidRatio# 'Humidity ratio [-]
    Dim vv# 'Kinematic viscosity [m²/s]
    Dim uu# 'Velocity [m/s]

    'Water properties
    rho = Water_Dens(T_K)
    vv = Water_KineVisc(T_K)

    'Flow properties
    darea = 0.25 * onePi * Diam_m * Diam_m 'Cross section area [m²]
    uu = FlowRate_m3s / darea 'nominal velocity [m/s]
    Re = uu * Diam_m / vv 'Reynolds number [-]
    ff = FrictionFactor(Re, Roughness_m / Diam_m) 'Darcy friction factor [-]
    Pdyn = 0.5 * rho * uu * uu 'Dynamic pressure [Pa]
    Water_PipeFriction = ff * Pdyn / Diam_m '[Pa/m]. Darcy-Weisbach equation: dP/L = f/D*Pdyn
End Function

Public Function FrictionFactor#(Re#, relRough#)
    'FrictionFactor [-] = Darcy friction factor for laminar or turbulent flow calculated with the Colebrook-White equation.
    'INPUTS:
    '- Re [-] = Reynolds number
    '- relRough [-] = Relative roughness = (RMS roughness, mm) / (Pipe internal diameter, mm)
    'NOTE: This algorithm has been tested against, and proved simpler and faster than two 3-log-call iterative methods: Serghides's
    'solution (Steffensen's method), Neta (desctibed by Praks & Brkic), and is more accurate than Serghides.
    'The reason this method appears to be quicker is that it uses Secant correction already after two logarithms.
    'Author: Peter.Schild@oslomet.no, 2020

    Const TransRe# = 2300
    Const pre1# = 0.868588963806504 '= 2*Log10(e) = 2/Ln(10) Used in solution of turbulent friction factor
    Const epsFA# = 0.000000001 'Absolute limit for convergence
    Dim iter& 'iteration number
    Dim x0# '= 1 / SQRT(ff[n-2])
    Dim x1# '= 1 / SQRT(ff[n-1])
    Dim x2# '= 1 / SQRT(ff[n])
    Dim f0# '=x1-x0
    Dim f1# '=x2-x1
    Dim pre2# 'Precalculated parameter = 2.51 / Re
    Dim pre3# 'Precalculated parameter = (Relative roughness) / 3.71

    If Re <= TransRe Then 'Laminar flow
        FrictionFactor = 64# / Re 'Darcy friction factor for laminar flow
    Else 'Turbulent flow
        pre2 = 2.51 / Re
        pre3 = relRough * 0.269541778975741
        'Quick estimate of initial value of x2=1/SQRT(f) without logarithms or powers. Fitted using Eureka in region 3E3<Re<1E9 and 0<e<0.05: R²=0.9954, MAE=0.145, MaxErr = 1.03
        x2 = 4.84319075863665 + 4.63779579749376 * Re / (11941706.1245753 + Re + 113734.583323483 * Re * relRough) + 4.1733649256718 * Re / (52647.4010173223 + Re + 770.82056958293 * Re * relRough) - 24.865138788607 * relRough
        For iter = 1 To 20 'Typically only 3-5 iterations needed
            If 2 < iter Then 'The first 2 passes fill x0 and x1
                If f1 = f0 Then Exit For
                x2 = x1 - f1 * (x1 - x0) / (f1 - f0) 'Secant correction, in a form that minimizes roundoff error
                If Abs(f1) < epsFA Then Exit For
            End If
            x0 = x1
            x1 = x2
            x2 = -pre1 * Log(pre2 * x2 + pre3)
            f0 = f1
            f1 = x2 - x1
            'Debug.Print iter, x2, f0, f1 '(just for testing)
        Next
        FrictionFactor = 1 / (x2 * x2)
    End If
End Function
