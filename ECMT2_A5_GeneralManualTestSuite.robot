*** Settings ***
Documentation           ECMT General Test
Test Setup              SuiteStartSetup
Test Teardown           SuiteStopTeardown       ${AllDataToOneExcel}    ${Test_Name}            ${DateTime}
Library                 RSC/CGT_Connect.py
Library                 RSC/BSP_KeywordLib.py
Library                 RSC/CAN_KeywordLib.py
Library                 RSC/CAN_KeywordLib.py    0                      WITH NAME               EXP_HSCAN
Library                 CAN_KeywordLib          WITH NAME               CGT_HSCA
Library                 RSC/EXL_KeywordLib.py
Library                 String
Library                 Collections
Library                 RSC/EXL_KeywordLib.py
Library                 DateTime
Library                 Dialogs
Metadata                Version                 V1.1-990425
Metadata                More Info               Writen by F.Naseri
Metadata                More Info               Modified for ECMT2 A5 Verification and Cleaned by Omid Seyedian
Metadata                Executed At             ${sResultFilePath}
Resource                RSC/ECMTGeneralKeywords.robot

#------------------------------------------------------------------------------------------------------------------
*** Variables ***
${CAN_FLAG}             ${0}
@{TMP}                  ${0}                    ${0}                    ${0}                    ${0}                    ${0}
@{DigitalInput}         ${0}                    ${0}                    ${0}                    ${0}                    ${0}
...                     ${0}                    ${0}                    ${0}                    ${0}
@{MU}                   ${0}                    ${0}                    ${0}                    ${0}                    ${0}
...                     ${0}                    ${0}                    ${0}                    ${0}                    ${0}
...                     ${0}                    ${0}                    ${0}                    ${0}                    ${0}                    ${0}
@{Freq_Outputs}         ${0}                    ${0}                    ${0}                    ${0}                    ${0}
...                     ${0}                    ${0}                    ${0}                    ${0}                    ${0}
...                     ${0}                    ${0}                    ${0}                    ${0}                    ${0}
...                     ${0}                    ${0}                    ${0}                    ${0}                    ${0}
${AllDataToOneExcel}
${Test_Name}
@{PullUpValueAN1}       ${0}                    # Real Value of +5V DC
@{PullUpValueAN2}       ${0}
@{PullUpValueAN3}       ${0}
@{PullUpValueAN4}       ${0}
@{MOSFET_Analog}        ${0}
@{SBCTLE_Analog}        ${0}
@{Chambr_Analog}        ${0}
@{WatPmp_Analog}        ${0}
${MOSFET_NTCTmp}        ${0}
${SBCTLE_NTCTmp}        ${0}
${Chambr_NTCTmp}        ${0}
${WatPmp_NTCTmp}        ${0}

#------------------------------------------------------------------------------------------------------------------
${MU_Version}           Proto-MU
${MC_Version}           I040
${HW_Version}           A4

#------------------------------------------------------------------------------------------------------------------
${Target_RPM}           770
${File_Path}            ${EXECDIR}/
${File_Name}            Report
${Excel_End}            X819
#${sFailedFilePath}     /home/eRI/TestReports/Fail
#${sResultFilePath}     /home/eRI/TestReports/Report
${sFailedFilePath}      ${EXECDIR}\\TestReports\\Fail
${sResultFilePath}      ${EXECDIR}\\TestReports\\
${MaxResponeseWaitingIteration}    ${1000}

#------------------------------------------------------------------------------------------------------------------
&{CMD_INDICES}          Excel_Crt=${100}        Excel_Wrt=${101}        CMD_TogglePort_A5=${102}    CMD_PWMoutput_A5=${103}
...                     CMD_ReportRequest_AnalogInputs_A5=${104}    CMD_ReportRequest_DigitalInputs_A5=${105}
...                     CMD_ReportRequest_FrequencyInputs_A5=${106}    CMD_ReportRequest_TLE8888_Diag=${107}
...                     CMD_work_mode=${108}    EXP_DinGet=${109}       EXP_DoSet=${110}        EXP_FoSet=${111}
...                     EXP_AoSet=${112}        CMD_ReadAllAnalogInputs=${113}    CMD_ReportRequest_ETC_Diag=${114}

#------------------------------------------------------------------------------------------------------------------
@{read_Version}                          ${3}    ${1}    ${2}
@{SaveEEPROMActive}                      ${4}    ${1}    ${5}    ${0}
@{SaveEEPROMDeactive}                    ${4}    ${1}    ${5}    ${1}
@{CMD_TogglePort}                        ${7}    ${1}    ${6}
@{CMD_PWMoutput}                         ${8}    ${1}    ${32}
@{CMD_ReadAllAnalogInputs}               ${3}    ${1}    ${67}
@{CMD_ReportRequest_AnalogInputs}        ${5}    ${1}    ${55}    ${0x10}
@{CMD_ReportRequest_DigitalInputs}       ${5}    ${1}    ${55}    ${0x20}
@{CMD_ReportRequest_FrequencyInputs}     ${5}    ${1}    ${55}    ${0x30}
@{CMD_ReportRequest_TLE8888_Diag}        ${5}    ${1}    ${55}    ${0x53}
@{CMD_ReportRequest_ETC_Diag}            ${5}    ${1}    ${55}    ${0x60}
@{CMD_ReportRequest_KP254}               ${5}    ${1}    ${55}    ${0x92}
@{CMD_ReportRequest_MCUTemperature}      ${5}    ${1}    ${55}    ${0x91}
@{CMD_ReportRequest_InternalStatus}      ${5}    ${1}    ${55}    ${0x00}
@{CMD_ReportRequest_MonitoringUnit}      ${5}    ${1}    ${55}    ${0x40}
@{CMD_work_mode}                         ${6}    ${1}    ${53}
@{CMD_REGWrite_TLE8888}                  ${5}    ${1}    ${64}
&{CMD_work_modes}       standby=${0}            idle=${1}               part_load=${2}          full_load=${3}          worst_case=${4}
&{CMD_TogglePort_A5}    CPPWM=${389}            LSD_EN=${-1}            SBC_DIS=${143}          ALT_CMD=${391}          LSHPWM_DOWN=${407}
...                     SOV1=${-1}              SOV2=${-1}              SOV3=${-1}              LSHPWM_UP=${406}
...                     IV_4=${402}             IV_3=${408}             IV_2=${393}             IV_1=${390}             IGC4=${126}             IGC3=${125}
...                     IGC2=${124}             IGC1=${123}             CLAMP1=${-1}
...                     CLAMP2=${-1}            CLAMP3=${-1}            CLAMP4=${-1}            GIV1=${-1}              GIV2=${-1}              GIV3=${-1}
...                     GIV4=${-1}              IVVTPWM=${401}          ESS=${-1}
...                     ETC_DIR=${118}          ETC_PWM=${117}          VLSU_PU=${179}          VLSD_PU=${193}          WDT=${-1}
...                     ETC_DIS=${109}          SPI_CLK=${102}          NCS_ETC=${107}          CS_MU=${105}
...                     SPI_MOSI=${104}         NCS_LSD=${-1}           MC_RST_REQ=${91}        TOOL0=${92}             CS_SBC=${-1}
...                     CS_KP254=${106}         CS_TLE=${142}           EVVTPWM=${404}          FAN_L=${398}
...                     HOT_LAMP=${397}         IGNEN=${-1}             MIL=${399}              INJEN=${-1}             FAN_H=${394}
...                     STST_RLY1=${-1}         THERMOSTAT=${383}       VAC_PUMP_RLY=${400}     RLY_ACCOUT=${396}
...                     RLY_EFP=${385}          STST_RLY2=${-1}         RLY_START=${403}        RLY_STST=${405}         RCL=${384}              ELE_WATER_PUMP=${386}      WGPWM=${395}                    ELE_OIL_PUMP=${410}
&{CMD_PWMoutput_A5}     LSHPWM_UP=${10}         LSHPWM_DOWN=${15}       IV1=${20}               IV2=${21}               IV3=${22}               IV4=${23}                  CPPWM=${30}                     IVVTPWM=${40}             EVVTPWM=${45}
...                     ALT_CMD=${60}           ESS=${-1}               GIV1=${-1}              GIV2=${-1}              GIV3=${-1}              GIV4=${-1}                 SOV1=${-1}                      SOV3=${-1}                CLAMP1=${-1}
...                     CLAMP2=${-1}            CLAMP3=${-1}            CLAMP4=${-1}            MOS3=${-1}              WGPWM=${120}            RCL=${130}
&{AnalogInputs_A5}      V_IGK_MIN=${0}          V_IGK=${0}              V_IGK_MAX=${0}          V_EL_MIN=${1}           V_EL=${1}               V_EL_MAX=${1}              BAP_OUT_MIN=${2}            BAP_OUT=${2}              BAP_OUT_MAX=${2}
...                     IGC_DIAG_MIN=${3}       IGC_DIAG=${3}           IGC_DIAG_MAX=${3}       FAN_DIAG_MIN=${4}       FAN_DIAG=${4}           FAN_DIAG_MAX=${4}          FTL_MIN=${5}                FTL=${5}                  FTL_MAX=${5}
...                     TCO_MIN=${6}            TCO=${6}                TCO_MAX=${6}            MAP_MIN=${7}            MAP=${7}                MAP_MAX=${7}               TIA_MIN=${8}                TIA=${8}                  TIA_MAX=${8}
...                     VLS_DOWN_MIN=${9}       VLS_DOWN=${9}           VLS_DOWN_MAX=${9}       VLS_UP_MIN=${10}        VLS_UP=${10}            VLS_UP_MAX=${10}           CRUISE_CTL_MIN=${11}        CRUISE_CTL=${11}          CRUISE_CTL_MAX=${11}
...                     PVS1_MIN=${12}          PVS1=${12}              PVS1_MAX=${12}          PVS2_MIN=${13}          PVS2=${13}              PVS2_MAX=${13}             TPS1_MIN=${14}              TPS1=${14}                TPS1_MAX=${14}
...                     TPS2_MIN=${15}          TPS2=${15}              TPS2_MAX=${15}          PUT_MIN=${16}           PUT=${16}               PUT_MAX=${16}              BRAKE_VACCUM_MIN=${17}      BRAKE_VACCUM=${17}        BRAKE_VACCUM_MAX=${17}
...                     MODE_SW_MIN=${18}       MODE_SW=${18}           MODE_SW_MAX=${18}       ST_REQ_MIN=${19}        ST_REQ=${19}            ST_REQ_MAX=${19}           CRK_DIAG_MIN=${20}          CRK_DIAG_SW=${20}         CRK_DIAG_MAX=${20}
...                     NEUTRAL_GEAR_SW_MIN=${21}     NEUTRAL_GEAR_SW=${21}           NEUTRAL_GEAR_SW_MAX=${21}    PRS_PSTE_MIN=${22}           PRS_PSTE=${22}             PRS_PSTE_MAX=${22}          KNK_MIN=${23}             KNK=${23}     KNK_MAX=${23}
&{DigitalInputs_A5}     ACIN_CurLev=${0}              ACIN_LevChang=${1}              ACC_PRS1_CurLev=${0}         ACC_PRS1_LevChang=${1}       ACC_PRS2_CurLev=${0}       ACC_PRS2_LevChang=${1}      IGK_STATUS_CurLev=${0}    IGK_STATUS_LevChang=${1}
...                     BRAKE_LIGHT_SW_CurLev=${0}    BRAKE_LIGHT_SW_LevChang=${1}    BRAKE_TEST_SW_CurLev=${0}    BRAKE_TEST_SW_LevChang=${1}  CLUTCH_SW2_CurLev=${0}     CLUTCH_SW2_LevChang=${1}    KICK_DOWN_SW_CurLev=${0}  KICK_DOWN_SW_LevChang=${1}
&{FrequencyInputs_A5}   CAM_HIGH_MIN=${0}             CAM_HIGH_MAX=${0}               CAM_LOW_MIN=${0}             CAM_LOW_MAX=${0}
...                     WSS_HIGH_MIN=${0}       WSS_HIGH_MAX=${0}           WSS_LOW_MIN=${0}         WSS_LOW_MAX=${0}
...                     ALT_MON_HIGH_MIN=${0}   ALT_MON_HIGH_MAX=${0}       ALT_MON_LOW_MIN=${0}     ALT_MON_LOW_MAX=${0}
...                     CAM_EX_HIGH_MIN=${0}    CAM_EX_MON_HIGH_MAX=${0}    CAM_EX_LOW_MIN=${0}      CAM_EX_LOW_MAX=${0}
...                     CRANK_HIGH_MIN=${0}     CRANK_MON_HIGH_MAX=${0}     CRANK_EX_LOW_MIN=${0}    CRANK_EX_LOW_MAX=${0}
&{EXP_DinGet}           ALL=${0}                MIL=${1}                    Electric_Oil_Pump=${2}   Electric_Water_Pump=${3}    FAN_LOW=${4}            FAN_HIGH=${5}          HS_RLY=${6}             LS_RLY=${7}
...                     EFP_OUT=${8}            VEL=${9}                    ACC_OUT=${10}           MAIN_PWR=${11}          BRAKE_VACCUM_PUMP=${12}    RCL=${13}              Electric_Thermostat=${14}    SOV1=${15}
...                     SOV2=${16}              SOV3=${17}
&{EXP_DoSet}            IGK=${0}                ACCPRS1=${1}            ACCPRS2=${2}            ACIN=${3}               PSTE=${4}               VEL=${5}               BTS=${6}                BLS=${7}
...                     CLUSWI=${8}             VBD=${9}                STRREQ=${10}            PV_KICK_DOWN_SW=${11}    FUEL_TYPE_SEL=${12}     NEUTRALGEARSW=${14}    CLUSWI1=${15}
&{EXP_FoSet}            RPM=${1}                VS=${2}                 CAMEX=${3}              Knock=${4}              ALT_MON=${5}
#RPM : cranck
#VS: WSS
&{EXP_AoSet}            FAN_Diag=${0}           TCO=${1}                TIA=${2}                   MAP=${3}                PVS=${4}                TPS=${5}               FTL=${6}                Cruise=${7}
...                     VLS_UP=${8}             VLS_Down=${9}           BrakeVac=${10}             PUT=${11}               MODE_SW=${12}           T_GAS_L=${13}          P_GAS_H=${14}           P_GAS_L=${15}
&{DiagResponse}         NoFailure=NoFailure     ShortToBAT=ShortToBAT   OpenLoad=OpenLoad          ShortToGround=ShortToGround
...                     OverTemperature=OverTemperature                 OverCurrent=OverCurrent    NoOverTemperature=NoOverTemperature
&{TLE8888_Diag_A5}      Diag0=${0}              Byte7=${0}              Byte6=${0}
...                     Byte5=${0}              Byte4=${0}
...                     Byte3=${0}              Byte2=${0}
...                     Byte1=${0}              Byte0=${0}
...                     IGN4=${DiagResponse.NoFailure}               IGN3=${DiagResponse.NoFailure}         #Byte7
...                     IGN2==${DiagResponse.NoFailure}              IGN1=${DiagResponse.NoFailure}         #Byte7
...                     RLY_EFP=${DiagResponse.NoFailure}            RLY_ACCOUT=${DiagResponse.NoFailure}   #Byte6 & Byte5
...                     RLY_FAN_HIGH=${DiagResponse.NoFailure}       RLY_STST=${DiagResponse.NoFailure}     #Byte6 & Byte5
...                     RLY_VAC_PUMP=${DiagResponse.NoFailure}       HOT_LAMP=${DiagResponse.NoFailure}     #Byte4
...                     MIL=${DiagResponse.NoFailure}                RLY_START=${DiagResponse.NoFailure}    #Byte4
...                     RLY_FAN_LOW=${DiagResponse.NoFailure}        ALT_CMD=${DiagResponse.NoFailure}      #Byte3
...                     ELE_THERMOSTAT=${DiagResponse.NoFailure}     IV3_CYL4=${DiagResponse.NoFailure}     #Byte3
...                     DO_ELE_OIL_PUMP=${DiagResponse.NoFailure}    DO_RCL=${DiagResponse.NoFailure}       #Byte2
...                     DO_EVVTPWM=${DiagResponse.NoFailure}         DO_WGPWM=${DiagResponse.NoFailure}     #Byte2
...                     DO_IVVTPWM=${DiagResponse.NoFailure}         LSHPWM_UP=${DiagResponse.NoFailure}    #Byte1
...                     LSHPWM_DN=${DiagResponse.NoFailure}          CPPWM=${DiagResponse.NoFailure}        #Byte1
...                     IV4_CYL2=${DiagResponse.NoFailure}           IV3_CYL4=${DiagResponse.NoFailure}     #Byte0
...                     IV2_CYL3=${DiagResponse.NoFailure}           IV1_CYL1=${DiagResponse.NoFailure}     #Byte0
&{ETCdiagResponse}      NoOverCurrent=NoOverCurrent   OverCurrent=OverCurrent    LoadShort=LoadShort
...                     isHigh=isHigh                 isLow=isLow    HWSC_LBIST_NotDone=HWSC_LBIST_NotDone
...                     HWSC_LBIST_Fail_Fail=HWSC_LBIST_Fail_Fail    HWSC_LBIST_Running_Pass=HWSC_LBIST_Running_Pass    HWSC_LBIST_Fail_Pass=HWSC_LBIST_Fail_Pass
...                     HWSC_LBIST_Pass_Pass=HWSC_LBIST_Pass_Pass    GlobalFailure=GlobalFailure            NoGlobalFailure=NoGlobalFailure
&{ETC_Diag_A5}          overcurrentMonitoring=${0}    statesResponse1=${0}
...                     statesResponse2=${0}          statesResponse3=${0}
...                     VDD_OV_UV=${0}
...                     OUT1_H=${ETCdiagResponse.NoOverCurrent}    OUT0_H=${ETCdiagResponse.NoOverCurrent}
...                     OUT1_L=${ETCdiagResponse.NoOverCurrent}    OUT0_L=${ETCdiagResponse.NoOverCurrent}
...                     NDIS=${ETCdiagResponse.isHigh}             DIS=${ETCdiagResponse.isLow}
...                     BRIDGE=${ETCdiagResponse.isHigh}           HWSC_LBIST=${ETCdiagResponse.HWSC_LBIST_NotDone}
...                     VPS_UV_REG=${ETCdiagResponse.isLow}        NGFAIL=${ETCdiagResponse.isHigh}
...                     ILIM_REG=${ETCdiagResponse.isLow}          VDD_OV_REG=${ETCdiagResponse.isLow}
...                     VDD_UV_REG=${ETCdiagResponse.isLow}        VPS_UV=${ETCdiagResponse.isLow}
...                     OTSDcnt=${ETCdiagResponse.isHigh}          OT_WARN=${ETCdiagResponse.isLow}
...                     OT_WARN_REG=${ETCdiagResponse.isLow}       NOTSD=${ETCdiagResponse.isHigh}
...                     NOTSD_REG=${ETCdiagResponse.isHigh}
&{KP254_A5}             PRESSURE=${0}               TEMPERATURE=${0}
...                     DIAG_H=${0}                 DIAG_L=${0}
@{MCresetTypes}         Power-On Reset              External Reset
...                     Loss of Lock Reset          Loss of Clock Reset
...                     Watchdog Timer Reset        Check stop Reset
...                     Software Watchdog Timer     Software System
@{MUresetTypes}         Power-On Reset              External Reset
...                     Internal Reset              MU Reset Out
...                     MC Reset Request
&{InternalStatus_A5}    CurrentRunningMode=UN-KNOWN    MClastReset=${0}
...                     MUlastReset=${0}               MUresetCount=${0}
...                     MUerrorCode=${0}               MUstate=${0}
${V_IGK_coef}           ${5.7667}
${V_EL_coef}            ${5.7667}
@{FIU_IP}               10.42.0.21
@{EXP_IP}               10.42.0.11
@{FIU_Stat}             SCG                     SCB
${FIU_Wait}             ${0.001}

*** Test Cases ***
#MyVElTest
#    # ${RET}    Toggle    1    2    100    50
#    # Sleep    1000
#    #Sleep    10
#    ${ECU_Version}          SendGiveDatafromECU     ${read_Version}
#    ${ECU_Version}          Run keyword If          ${ECU_Version} == ${False}    Log To Console          Read Version Failed
#    ...                     ELSE                    Convert to ASCII        ${ECU_Version}[1][4:]
#    Run Keyword If          ${ECU_Version} != ${None}    Log To Console          ${ECU_Version}
#    ${Version_part1}        Evaluate                ''.join(str(e) for e in ${ECU_Version}[0:5])
#    ${Version_part2}        Evaluate                ''.join(str(e) for e in ${ECU_Version}[5:8])
#    ${Version_part3}        Evaluate                ''.join(str(e) for e in ${ECU_Version}[8:12])
#    Log To Console          ${Version_part1} ${Version_part2} ${Version_part3}
#    # ${RET}    Toggle    2    1    100    10
#    #Sleep    20
#    Start Work mode    ${3}    ${1}    ${1}
#    ${DateTime}                Get Current Date        local                   0                       timestamp
#    
#    FOR    ${loopCounter}    IN RANGE    ${1}    ${1000}
#        Log To Console    ${loopCounter}
#        Sleep    4
#        Do Set    ${5}    0
#        #Sleep    1ms
#        Do Set    ${5}    1        # 5 == VEL
#        Log To Console    ${loopCounter}
#    END
    
EnduranceTest
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/HW+Validation+%28A5%29+-+Electrical
    FinSelMux_Set
    ${Test_Name}               Set Variable            Endurance Test          #first word of output excel files
    ${DateTime}                Get Current Date        local                   0                       timestamp               ${True}                #second word
    ${DateTime}                Replace String          ${DateTime}             :                       _                       2
    ${V_IGK}                   Set Variable            ${13.5}
    ${V_EL}                    Set Variable            ${13.5}
    ${FAN_Diag_AnalogValue}    Set Variable            ${2.5}
    ${TCO_AnalogValue}         Set Variable            ${1.5}
    ${TIA_AnalogValue}         Set Variable            ${1.5}
    ${MAP_AnalogValue}         Set Variable            ${1.5}
    ${PVS_AnalogValue}         Set Variable            ${1.5}
    ${TPS_AnalogValue}         Set Variable            ${1.5}
    ${FTL_AnalogValue}         Set Variable            ${1.5}
    ${Cruise_AnalogValue}      Set Variable            ${1.5}
    ${VLS_UP_AnalogValue}      Set Variable            ${1.5}
    ${VLS_Down_AnalogValue}    Set Variable            ${1.5}
    ${BrakeVac_AnalogValue}    Set Variable            ${1.5}
    ${PUT_AnalogValue}         Set Variable            ${1.5}
    ${MODE_SW_AnalogValue}     Set Variable            ${1.5}        # MODE_SW has a volatge divider. 4.8*10/32 = 1.5
    ${T_GAS_L_AnalogValue}     Set Variable            ${1.5}
    ${P_GAS_H_AnalogValue}     Set Variable            ${1.5}
    ${P_GAS_L_AnalogValue}     Set Variable            ${1.5}
    sleep                   1
    ${ECU_Version}          SendGiveDatafromECU     ${read_Version}
    ${ECU_Version}          Run keyword If          ${ECU_Version} == ${False}    Log To Console          Read Version Failed
    ...                     ELSE                    Convert to ASCII        ${ECU_Version}[1][4:]
    Run Keyword If          ${ECU_Version} != ${None}    Log To Console          ${ECU_Version}
    ${Version_part1}        Evaluate                ''.join(str(e) for e in ${ECU_Version}[0:5])
    ${Version_part2}        Evaluate                ''.join(str(e) for e in ${ECU_Version}[5:8])
    ${Version_part3}        Evaluate                ''.join(str(e) for e in ${ECU_Version}[8:12])
    Log To Console          ${Version_part1} ${Version_part2} ${Version_part3}
    
    # ${EEPROM_Resp}          SendGiveDatafromECU     ${SaveEEPROMActive}
    # Run Keyword If          ${EEPROM_Resp}[2] != ${0xA0}    Log To Console          SaveEEPROMActiveError
    
    ${Mode}                 Get Selection From User    Please choose the test's work mode:    standby                 idle                    part_load              full_load               worst_case
    ${TestName}             Get Value From User     Please specify test's name:    Endurance Test
    ${LabelRow1_OfExcel}    Create List             Test Name               ${TestName}
    ${LabelRow2_OfExcel}    Create List             Test Mode               ${Mode}
    ${LabelRow3_OfExcel}    Create List             Date & Time             ${DateTime}
    ${LabelRow4_OfExcel}    Create List             EXP. SW Ver.            Ver. IV
    ${LabelRow5_OfExcel}    Create List             Excel Part              Part1
    ${LabelRow6_OfExcel}    Create List             MU SW Ver.              ${Version_part1} ${Version_part2} ${Version_part3}
    ${LabelRow7_OfExcel}    Create List             Test Done By            Omid Seyedian
    ${LabelRow8_OfExcel}    Create List             Parameters
    ...                     V_IGK                   V_IGK                   V_IGK
    ...                     V_EL                    V_EL                    V_EL
    ...                     BAP_OUT                 BAP_OUT                 BAP_OUT
    ...                     IGC_DIAG                IGC_DIAG                IGC_DIAG
    ...                     FAN_DIAG                FAN_DIAG                FAN_DIAG
    ...                     FTL                     FTL                     FTL
    ...                     TCO                     TCO                     TCO
    ...                     MAP                     MAP                     MAP
    ...                     TIA                     TIA                     TIA
    ...                     VLS_DOWN                VLS_DOWN                VLS_DOWN
    ...                     VLS_UP                  VLS_UP                  VLS_UP
    ...                     CRUISE_CTL              CRUISE_CTL              CRUISE_CTL
    ...                     PVS1                    PVS1                    PVS1
    ...                     PVS2                    PVS2                    PVS2
    ...                     TPS1                    TPS1                    TPS1
    ...                     TPS2                    TPS2                    TPS2
    ...                     PUT                     PUT                     PUT
    ...                     BRAKE_VACCUM            BRAKE_VACCUM            BRAKE_VACCUM
    ...                     MODE_SW                 MODE_SW                 MODE_SW
    ...                     ST_REQ                  ST_REQ                  ST_REQ
    ...                     CRK_DIAG                CRK_DIAG                CRK_DIAG
    ...                     NEUTRAL_GEAR_SW         NEUTRAL_GEAR_SW         NEUTRAL_GEAR_SW
    ...                     PRS_PSTE                PRS_PSTE                PRS_PSTE
    ...                     KNK                     KNK                     KNK
    ...                     ACIN                    ACIN                    ACIN
    ...                     ACC_PRS1                ACC_PRS1                ACC_PRS1
    ...                     ACC_PRS2                ACC_PRS2                ACC_PRS2
    ...                     IGK_STATUS              IGK_STATUS              IGK_STATUS
    ...                     CLUTCH_SW               CLUTCH_SW               CLUTCH_SW
    ...                     BRAKE_LIGHT_SW          BRAKE_LIGHT_SW          BRAKE_LIGHT_SW
    ...                     BRAKE_TEST_SW           BRAKE_TEST_SW           BRAKE_TEST_SW
    ...                     CLUTCH_SW2              CLUTCH_SW2              CLUTCH_SW2
    ...                     CAM_HIGH                CAM_HIGH                CAM_HIGH
    ...                     CAM_LOW                 CAM_LOW                 CAM_LO
    ...                     WSS_HIGH                WSS_HIGH                WSS_HIGH
    ...                     WSS_LOW                 WSS_LOW                 WSS_LOW
    ...                     ALT_MON_HIGH            ALT_MON_HIGH            ALT_MON_HIGH
    ...                     ALT_MON_LOW             ALT_MON_LOW             ALT_MON_LOW
    ...                     CAM_EX_HIGH             CAM_EX_HIGH             CAM_EX_HIGH
    ...                     CAM_EX_LOW              CAM_EX_LOW              CAM_EX_LOW
    ...                     CRANK_HIGH              CRANK_HIGH              CRANK_HIGH
    ...                     CRANK_LOW               CRANK_LOW               CRANK_LOW
    ...                     IGNDIAG                 BRIDIAG1                BRIDIAG0
    ...                     OUTDIAG4                OUTDIAG3                OUTDIAG2
    ...                     OUTDIAG1                OUTDIAG0
    ...                     ETC_OVERCURRENT         ETC_STATE_RESP1         ETC_STATE_RESP2         ETC_STATE_RESP3
    ...                     ETC_VDD_OV_UV           ETC_REG                 ETC_REG                 ETC_REG
    ...                     KP254_PRESSURE          KP254_TEMP              KP254_DIAG_H            KP254_DIAG_L
    ...                     MOSFET_Tmp              SBCTLE_Tmp              Chambr_Tmp              WatPmp_Temp
    ...                     InternalStatus_A5.MClastReset     InternalStatus_A5.MUlastReset         InternalStatus_A5.MUresetCount
    ...                     InternalStatus_A5.MUerrorCode     InternalStatus_A5.MUstate             MCU Temperature
    ...                     MU                      MU                      MU                      MU
    ...                     MU                      MU                      MU                      MU
    ...                     MU                      MU                      MU                      MU
    ...                     MU                      MU                      MU                      MU
    ...                     Frequency_Output        Frequency_Output        Frequency_Output        Frequency_Output
    ...                     Frequency_Output        Frequency_Output        Frequency_Output        Frequency_Output
    ...                     Frequency_Output        Frequency_Output        Frequency_Output        Frequency_Output
    ...                     Frequency_Output        Frequency_Output        Frequency_Output        Frequency_Output
    ...                     Frequency_Output        Frequency_Output        Frequency_Output        Frequency_Output
    ...                     Renegade!               Renegade!               Renegade!
    ...                     KICK_DOWN_SW            KICK_DOWN_SW            KICK_DOWN_SW

    ${LabelRow9_OfExcel}    Create List             Sample Time
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX   
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   CUR                     CHNG
    ...                     VALUE                   CUR                     CHNG
    ...                     VALUE                   CUR                     CHNG
    ...                     VALUE                   CUR                     CHNG
    ...                     VALUE                   CUR                     CHNG
    ...                     VALUE                   CUR                     CHNG
    ...                     VALUE                   CUR                     CHNG
    ...                     VALUE                   CUR                     CHNG
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX  
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     VALUE                   MIN                     MAX
    ...                     TLE_REG                 TLE_REG                 TLE_REG                 TLE_REG
    ...                     TLE_REG                 TLE_REG                 TLE_REG                 TLE_REG
    ...                     ETC_REG                 ETC_REG                 ETC_REG                 ETC_REG
    ...                     ETC_REG                 ETC_REG                 ETC_REG                 ETC_REG
    ...                     VALUE                   VALUE                   VALUE                   VALUE
    ...                     VALUE                   VALUE                   VALUE                   VALUE
    ...                     VALUE                   VALUE                   VALUE                   VALUE
    ...                     VALUE                   VALUE
    ...                     Byte1                   Byte2                   Byte3                   Byte4
    ...                     Byte5                   Byte6                   Byte7                   Byte8
    ...                     Byte9                   Byte10                  Byte11                  Byte12
    ...                     Byte13                  Byte14                  Byte15                  Byte16
    ...                     ALT_CMD                 CP                      EVVT
    ...                     GIV1                    GIV2                    GIV3                    GIV4
    ...                     IGC1                    IGC2                    IGC3                    IGC4
    ...                     INJ1                    INJ2                    INJ3                    INJ4
    ...                     IVVT                    LSHDOWN                 LSHUP                   WG_NEG
    ...                     WG                      Diag0                   MU_Mode                 MC_ExRst
    ...                     VALUE                   CUR                     CHNG
            
    ${ExcelFileNumber}      Set Variable            ${0}
    ${Excel_row}            Set Variable            ${0}
    ${MaxRowsInOneFile}     Set Variable            ${25}
    ${OveralRepeat}         Set Variable            ${25}
    ${ExcelParams}          Create List
    ${rowContent}           Create List
    ${AllDataToOneExcel}    Create List             ${LabelRow1_OfExcel}
    Append To List          ${AllDataToOneExcel}    ${LabelRow2_OfExcel}
    Append To List          ${AllDataToOneExcel}    ${LabelRow3_OfExcel}
    Append To List          ${AllDataToOneExcel}    ${LabelRow4_OfExcel}
    Append To List          ${AllDataToOneExcel}    ${LabelRow5_OfExcel}
    Append To List          ${AllDataToOneExcel}    ${LabelRow6_OfExcel}
    Append To List          ${AllDataToOneExcel}    ${LabelRow7_OfExcel}
    Append To List          ${AllDataToOneExcel}    ${LabelRow8_OfExcel}
    Append To List          ${AllDataToOneExcel}    ${LabelRow9_OfExcel}
    # EXPanalogWrite        ${EXP_AoSet.FAN_Diag}    ${FAN_Diag_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.TCO}        ${TCO_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.TIA}        ${TIA_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.MAP}        ${MAP_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.PVS}        ${PVS_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.TPS}        ${TPS_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.FTL}        ${FTL_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.Cruise}     ${Cruise_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.VLS_UP}     ${VLS_UP_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.VLS_Down}    ${VLS_Down_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.BrakeVac}    ${BrakeVac_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.PUT}        ${PUT_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.MODE_SW}    ${MODE_SW_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.T_GAS_L}    ${T_GAS_L_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.P_GAS_H}    ${P_GAS_H_AnalogValue}
    EXPanalogWrite          ${EXP_AoSet.P_GAS_L}    ${P_GAS_L_AnalogValue}
    Start Work mode         ${CMD_work_modes.${Mode}}    ${1}                    ${1}
    sleep                   1
    # # CAM
    # Fo Set                ${???}                  ${100}                  ${50}
    # WSS
    ${ret}                  Fo Set                  ${2}                    ${100}                  ${30}
    # Log To Console        ${ret}
    # ALT_MON
    ${ret}                  Fo Set                  ${5}                    ${250}                  ${10}
    # Log To Console        ${ret}
    # CAM_EX
    ${ret}                  Fo Set                  ${3}                    ${100}                  ${60}
    # Log To Console        ${ret}
    # KNK
    ${ret}                  Fo Set                  ${4}                    ${20000}                ${30}
    # CRNK
    ${ret}                  Fo Set                  ${1}                    ${6000}                 ${50}                   ${1}
    # Log To Console        ${ret}
    # ACIN
    ${Ret_List}             Do Set                  ${3}                    ${1}
    # ACCPRS1
    ${Ret_List}             Do Set                  ${1}                    ${1}
    # ACCPRS2
    ${Ret_List}             Do Set                  ${2}                    ${1}
    # IGK
    # ${Ret_List}           Do Set                  ${0}                    ${1}
    # CLUSWI
    ${Ret_List}             Do Set                  ${8}                    ${1}
    # BLS
    ${Ret_List}             Do Set                  ${7}                    ${1}
    # BTS
    ${Ret_List}             Do Set                  ${6}                    ${1}
    # CLUSWI1
    ${Ret_List}             Do Set                  ${15}                   ${1}
    FOR                     ${loopCounter}          IN RANGE                ${1}                    ${OveralRepeat}
        ${Omid}                 Cgt Hw Ver Get
        ${newStart}             Evaluate                ${loopCounter} % ${MaxRowsInOneFile}
        ${DateTime2}                Get Current Date        local                   0                       timestamp               ${True}                #second word
        ${DateTime2}                Replace String          ${DateTime2}             :                       _                 
        Log To Console          ${loopCounter}) ${DateTime2} 
        ${SampleTime}           Get Current Date
        ${SampleTime}           Get Substring           ${SampleTime}           11
        Run Keyword If          ${loopCounter} % 6 > 2    ReSet_DigitalInputs
        Run Keyword If          ${loopCounter} % 6 <= 2    Set_DigitalInputs
        Run Keyword If          ${loopCounter} % 8 > 3    Do Set                  ${10}                   ${1}
        Run Keyword If          ${loopCounter} % 8 > 3    Do Set                  ${11}                   ${1}
        Run Keyword If          ${loopCounter} % 8 > 3    Do Set                  ${14}                   ${1}
        Run Keyword If          ${loopCounter} % 8 > 3    Do Set                  ${4}                    ${1}
        Run Keyword If          ${loopCounter} % 8 <= 3    Do Set                  ${10}                   ${0}
        Run Keyword If          ${loopCounter} % 8 <= 3    Do Set                  ${11}                   ${0}
        Run Keyword If          ${loopCounter} % 8 <= 3    Do Set                  ${14}                   ${0}
        Run Keyword If          ${loopCounter} % 8 <= 3    Do Set                  ${4}                    ${0}
        Run Keyword If          ${loopCounter} % 5 == 0    ReportRequest_Temperature
        ReportRequest_AnalogInputs    ${1}
        ${argus}                Evaluate                ${loopCounter} % 8
        AoSel_AnalogInputs      ${loopCounter}
        ReportRequest_DigitalInputs    ${1}
        ReportRequest_FrequencyInputs    ${1}
        ReportRequest_TLE8888_Diag    ${0}
        ReportRequest_ETC_Diag    ${0}
        ReportRequest_InternalStatus    ${0}
        ReportRequest_KP254     ${1}
        ReportRequest_MonitoringUnit    ${0}
        ReportRequest_MCUTemperature    ${0}
        Run Keyword If          ${loopCounter} % 8 <= 3    ReportRequest_FrequencyOutputs_1
        Run Keyword If          ${loopCounter} % 8 <= 3    FinSelMux_ReSet
        Run Keyword If          ${loopCounter} % 8 > 3    ReportRequest_FrequencyOutputs_2
        Run Keyword If          ${loopCounter} % 8 > 3    FinSelMux_Set
        
        ${rowContent}           Create List                ${SampleTime}
        ...                     ${V_IGK}                   ${AnalogInputs_A5.V_IGK_MIN}               ${AnalogInputs_A5.V_IGK_MAX}
        ...                     ${V_EL}                    ${AnalogInputs_A5.V_EL_MIN}                ${AnalogInputs_A5.V_EL_MAX}
        ...                     ??                         ${AnalogInputs_A5.BAP_OUT_MIN}             ${AnalogInputs_A5.BAP_OUT_MAX}
        ...                     ${0.330}                   ${AnalogInputs_A5.IGC_DIAG_MIN}            ${AnalogInputs_A5.IGC_DIAG_MAX}
        ...                     ??                         ${AnalogInputs_A5.FAN_DIAG_MIN}            ${AnalogInputs_A5.FAN_DIAG_MAX}
        ...                     ${FTL_AnalogValue}         ${AnalogInputs_A5.FTL_MIN}                 ${AnalogInputs_A5.FTL_MAX}
        ...                     ${TCO_AnalogValue}         ${AnalogInputs_A5.TCO_MIN}                 ${AnalogInputs_A5.TCO_MAX}
        ...                     ${MAP_AnalogValue}         ${AnalogInputs_A5.MAP_MIN}                 ${AnalogInputs_A5.MAP_MAX}
        ...                     ${TIA_AnalogValue}         ${AnalogInputs_A5.TIA_MIN}                 ${AnalogInputs_A5.TIA_MAX}
        ...                     ${VLS_Down_AnalogValue}    ${AnalogInputs_A5.VLS_DOWN_MIN}            ${AnalogInputs_A5.VLS_DOWN_MAX}
        ...                     ${VLS_UP_AnalogValue}      ${AnalogInputs_A5.VLS_UP_MIN}              ${AnalogInputs_A5.VLS_UP_MAX}
        ...                     ${Cruise_AnalogValue}      ${AnalogInputs_A5.CRUISE_CTL_MIN}          ${AnalogInputs_A5.CRUISE_CTL_MAX}
        ...                     ${PVS_AnalogValue}         ${AnalogInputs_A5.PVS1_MIN}                ${AnalogInputs_A5.PVS1_MAX}
        ...                     ${0.75}                    ${AnalogInputs_A5.PVS2_MIN}                ${AnalogInputs_A5.PVS2_MAX}
        ...                     ${TPS_AnalogValue}         ${AnalogInputs_A5.TPS1_MIN}                ${AnalogInputs_A5.TPS1_MAX}
        ...                     ${3.5}                     ${AnalogInputs_A5.TPS2_MIN}                ${AnalogInputs_A5.TPS2_MAX}
        ...                     ${PUT_AnalogValue}         ${AnalogInputs_A5.PUT_MIN}                 ${AnalogInputs_A5.PUT_MAX}
        ...                     ${BrakeVac_AnalogValue}    ${AnalogInputs_A5.BRAKE_VACCUM_MIN}        ${AnalogInputs_A5.BRAKE_VACCUM_MAX}
        ...                     ${MODE_SW_AnalogValue}     ${AnalogInputs_A5.MODE_SW_MIN}             ${AnalogInputs_A5.MODE_SW_MAX}
        ...                     ??                         ${AnalogInputs_A5.ST_REQ_MIN}              ${AnalogInputs_A5.ST_REQ_MAX}
        ...                     ${2}                       ${AnalogInputs_A5.CRK_DIAG_MIN}            ${AnalogInputs_A5.CRK_DIAG_MAX}
        ...                     ??                         ${AnalogInputs_A5.NEUTRAL_GEAR_SW_MIN}     ${AnalogInputs_A5.NEUTRAL_GEAR_SW_MAX}
        ...                     ??                         ${AnalogInputs_A5.PRS_PSTE_MIN}            ${AnalogInputs_A5.PRS_PSTE_MAX}
        ...                     ${2.46}                    ${AnalogInputs_A5.KNK_MIN}                 ${AnalogInputs_A5.KNK_MAX}
        ...                     ${DigitalInput}[0]         ${DigitalInputs_A5.ACIN_CurLev}            ${DigitalInputs_A5.ACIN_LevChang}
        ...                     ${DigitalInput}[1]         ${DigitalInputs_A5.ACC_PRS1_CurLev}        ${DigitalInputs_A5.ACC_PRS1_LevChang}
        ...                     ${DigitalInput}[2]         ${DigitalInputs_A5.ACC_PRS2_CurLev}        ${DigitalInputs_A5.ACC_PRS2_LevChang}
        ...                     ${DigitalInput}[3]         ${DigitalInputs_A5.IGK_STATUS_CurLev}      ${DigitalInputs_A5.IGK_STATUS_LevChang}
        ...                     ${DigitalInput}[4]         ${DigitalInputs_A5.CLUTCH_SW_CurLev}       ${DigitalInputs_A5.CLUTCH_SW_LevChang}
        ...                     ${DigitalInput}[5]         ${DigitalInputs_A5.BRAKE_LIGHT_SW_CurLev}  ${DigitalInputs_A5.BRAKE_LIGHT_SW_LevChang}
        ...                     ${DigitalInput}[6]         ${DigitalInputs_A5.BRAKE_TEST_SW_CurLev}   ${DigitalInputs_A5.BRAKE_TEST_SW_LevChang}
        ...                     ${DigitalInput}[7]         ${DigitalInputs_A5.CLUTCH_SW2_CurLev}      ${DigitalInputs_A5.CLUTCH_SW2_LevChang}
        ...                     ??                         ${FrequencyInputs_A5.CAM_HIGH_MIN}         ${FrequencyInputs_A5.CAM_HIGH_MAX}
        ...                     ??                         ${FrequencyInputs_A5.CAM_LOW_MIN}          ${FrequencyInputs_A5.CAM_LOW_MAX}
        ...                     7000                       ${FrequencyInputs_A5.WSS_HIGH_MIN}         ${FrequencyInputs_A5.WSS_HIGH_MAX}
        ...                     3000                       ${FrequencyInputs_A5.WSS_LOW_MIN}          ${FrequencyInputs_A5.WSS_LOW_MAX}
        ...                     400                        ${FrequencyInputs_A5.ALT_MON_HIGH_MIN}     ${FrequencyInputs_A5.ALT_MON_HIGH_MAX}
        ...                     3600                       ${FrequencyInputs_A5.ALT_MON_LOW_MIN}      ${FrequencyInputs_A5.ALT_MON_LOW_MAX}
        ...                     ??                         ${FrequencyInputs_A5.CAM_EX_HIGH_MIN}      ${FrequencyInputs_A5.CAM_EX_HIGH_MAX}
        ...                     ??                         ${FrequencyInputs_A5.CAM_EX_LOW_MIN}       ${FrequencyInputs_A5.CAM_EX_LOW_MAX}
        ...                     83.3                       ${FrequencyInputs_A5.CRANK_HIGH_MIN}       ${FrequencyInputs_A5.CRANK_HIGH_MAX}
        ...                     83.3                       ${FrequencyInputs_A5.CRANK_LOW_MIN}        ${FrequencyInputs_A5.CRANK_LOW_MAX}
        ...                     ${TLE8888_Diag_A5.Byte7}   ${TLE8888_Diag_A5.Byte6}                   ${TLE8888_Diag_A5.Byte5}
        ...                     ${TLE8888_Diag_A5.Byte4}   ${TLE8888_Diag_A5.Byte3}                   ${TLE8888_Diag_A5.Byte2}
        ...                     ${TLE8888_Diag_A5.Byte1}   ${TLE8888_Diag_A5.Byte0} 
        ...                     ${ETC_Diag_A5.overcurrentMonitoring}
        ...                     ${ETC_Diag_A5.statesResponse1}    ${ETC_Diag_A5.statesResponse2}
        ...                     ${ETC_Diag_A5.statesResponse3}    ${ETC_Diag_A5.VDD_OV_UV}
        ...                     ${0}                    ${0}                    ${0}
        ...                     ${KP254_A5.PRESSURE}    ${KP254_A5.TEMPERATURE}
        ...                     ${KP254_A5.DIAG_H}      ${KP254_A5.DIAG_L}
        ...                     ${TMP}[0]               ${TMP}[1]               ${TMP}[2]               ${TMP}[3]
        ...                     ${InternalStatus_A5.MClastReset}   ${InternalStatus_A5.MUlastReset}
        ...                     ${InternalStatus_A5.MUresetCount}  ${InternalStatus_A5.MUerrorCode}  ${InternalStatus_A5.MUstate}
        ...                     ${TMP}[4]               ${MU}[0]                ${MU}[1]             ${MU}[2]
        ...                     ${MU}[3]                ${MU}[4]                ${MU}[5]             ${MU}[6]
        ...                     ${MU}[7]                ${MU}[8]                ${MU}[9]             ${MU}[10]
        ...                     ${MU}[11]               ${MU}[12]               ${MU}[13]            ${MU}[14]
        ...                     ${MU}[15]
        ...                     ${Freq_Outputs}[0]      ${Freq_Outputs}[1]      ${Freq_Outputs}[2]      ${Freq_Outputs}[3]
        ...                     ${Freq_Outputs}[4]      ${Freq_Outputs}[5]      ${Freq_Outputs}[6]      ${Freq_Outputs}[7]
        ...                     ${Freq_Outputs}[8]      ${Freq_Outputs}[9]      ${Freq_Outputs}[10]     ${Freq_Outputs}[11]
        ...                     ${Freq_Outputs}[12]     ${Freq_Outputs}[13]     ${Freq_Outputs}[14]     ${Freq_Outputs}[15]
        ...                     ${Freq_Outputs}[16]     ${Freq_Outputs}[17]     ${Freq_Outputs}[18]     ${Freq_Outputs}[19]
        ...                     ${TLE8888_Diag_A5.Diag0}   ${InternalStatus_A5.CurrentRunningMode}    ${InternalStatus_A5.MCExRstCount}
        ...                     ${DigitalInput}[8]         ${DigitalInputs_A5.KICK_DOWN_SW_CurLev}    ${DigitalInputs_A5.KICK_DOWN_SW_LevChang}
        
        Run Keyword If          ${CAN_FLAG}==${1}       Set List Value          ${rowContent}           1                       CAN Failure!
        Set Global Variable     ${CAN_FLAG}             ${0}
        Run Keyword If          ${loopCounter} > 5      Append To List          ${AllDataToOneExcel}    ${rowContent}
        ${rowContent}           Set Variable            ${null}
        
        ${ExcelFileNumber}      Run Keyword If          ${newStart} == 0        Evaluate                ${ExcelFileNumber} + ${1}
        ...                     ELSE                    Set Variable            ${ExcelFileNumber}
        ${ExcelFileName}        Run Keyword If          ${newStart} == 0        Catenate                ${Test_Name}            ${DateTime}            part                    ${ExcelFileNumber}
        ${ExcelParams}          Run Keyword If          ${newStart} == 0        Create Excel File       ${sResultFilePath}      ${ExcelFileName}       ${Test_Name}
        ...                     ELSE                    Set Variable            ${ExcelParams}
        ${Ret_Val}              Run Keyword If          ${newStart} == 0        Write Matrix To Row     ${sResultFilePath}      ${ExcelFileName}       ${Test_Name}            ${AllDataToOneExcel}    row=${1}                col=${1}
        ${AllDataToOneExcel}    Run Keyword If          ${newStart} == 0        Set Variable            ${null}
        ...                     ELSE                    Set Variable            ${AllDataToOneExcel}
        ${Excel_PartNumber}     Evaluate                ${ExcelFileNumber} + 1
        ${LabelRow5_OfExcel}    Set Variable            Excel Part              Part${Excel_PartNumber}
        ${AllDataToOneExcel}    Run Keyword If          ${newStart} == 0        Create List             ${LabelRow1OfExcel}
        ...     ${LabelRow2OfExcel}    ${LabelRow3OfExcel}     ${LabelRow4OfExcel}     ${LabelRow5OfExcel}     ${LabelRow6OfExcel}
        ...    ${LabelRow7OfExcel}    ${LabelRow8OfExcel}    ${LabelRow9OfExcel}
        ...                     ELSE                    Set Variable            ${AllDataToOneExcel}
    END

*** Keywords ***
ReportRequest_Temperature
    # This test case reports the temperature of different sections in the ECU and the chamber
    # Please Write down the Real +5V Values (From The Test Case: Pull Up Value Catcher) Below:
    ${PullUpValueAN1}       Set Variable            4.8979                  # These are the real values of the nominal +5V DC
    ${PullUpValueAN2}       Set Variable            4.8985
    ${PullUpValueAN3}       Set Variable            4.8988
    ${PullUpValueAN4}       Set Variable            4.8985
    ${Sample}               Set Variable            5                       # No. of temperature measurement
    ${dispose}              Set Variable            1                       # No. of measurements disposed
    # MOSFET -------------------------------------------------------------------------------------------------------
    FOR                     ${i}                    IN RANGE                0                       ${Sample}               # Polulating the iterration
        ${ret}                  Ain Get                 3                       # ADC
        ${analog}               Convert To Number       ${ret}[1]               # Pure analog measurement
        Append To List          ${MOSFET_Analog}        ${analog}
    END
    Sort List               ${MOSFET_Analog}
    # Sorting, Discarding absurdities and averaging, for more accurate and meaningful measurement
    ${MOSFET_Analog}        Evaluate                round(numpy.mean(${MOSFET_Analog}[${dispose}:-${dispose}]),4)
    #---------------------------------------------------------------------------------------------------------------
    # SBCTLE -------------------------------------------------------------------------------------------------------
    FOR                     ${i}                    IN RANGE                0                       ${Sample}               # Polulating the iterration
        ${ret}                  Ain Get                 4                       # ADC
        ${analog}               Convert To Number       ${ret}[1]               # Pure analog measurement
        Append To List          ${SBCTLE_Analog}        ${analog}
    END
    Sort List               ${SBCTLE_Analog}
    # Sorting, Discarding absurdities and averaging, for more accurate and meaningful measurement
    ${SBCTLE_Analog}        Evaluate                round(numpy.mean(${SBCTLE_Analog}[${dispose}:-${dispose}]),4)
    #---------------------------------------------------------------------------------------------------------------
    # Chambr -------------------------------------------------------------------------------------------------------
    FOR                     ${i}                    IN RANGE                0                       ${Sample}               # Polulating the iterration
        ${ret}                  Ain Get                 5                       # ADC
        ${analog}               Convert To Number       ${ret}[1]               # Pure analog measurement
        Append To List          ${Chambr_Analog}        ${analog}
    END
    Sort List               ${Chambr_Analog}
    # Sorting, Discarding absurdities and averaging, for more accurate and meaningful measurement
    ${Chambr_Analog}        Evaluate                round(numpy.mean(${Chambr_Analog}[${dispose}:-${dispose}]),4)
    #---------------------------------------------------------------------------------------------------------------
    # WatPmp -------------------------------------------------------------------------------------------------------
    FOR                     ${i}                    IN RANGE                0                       ${Sample}               # Polulating the iterration
        ${ret}                  Ain Get                 6                       # ADC
        ${analog}               Convert To Number       ${ret}[1]               # Pure analog measurement
        Append To List          ${WatPmp_Analog}        ${analog}
    END
    Sort List               ${WatPmp_Analog}
    # Sorting, Discarding absurdities and averaging, for more accurate and meaningful measurement
    ${WatPmp_Analog}        Evaluate                round(numpy.mean(${WatPmp_Analog}[${dispose}:-${dispose}]),4)
    #---------------------------------------------------------------------------------------------------------------
    # The following codes are calculating the resistance value of the NTCs
    # The absurdities (devide by 0) will rule out
    ${MOSFET_NTCRes}        Run Keyword If          abs(${PullUpValueAN1} - ${MOSFET_Analog}) != 0
    ...                     Evaluate                round((${MOSFET_Analog} * 820) / abs(${PullUpValueAN1} - ${MOSFET_Analog})) + 1
    ...                     ELSE                    Set Variable            10000
    ${SBCTLE_NTCRes}        Run Keyword If          abs(${PullUpValueAN2} - ${SBCTLE_Analog}) != 0
    ...                     Evaluate                round((${SBCTLE_Analog} * 820) / abs(${PullUpValueAN2} - ${SBCTLE_Analog})) + 1
    ...                     ELSE                    Set Variable            10000
    ${Chambr_NTCRes}        Run Keyword If          abs(${PullUpValueAN3} - ${Chambr_Analog}) != 0
    ...                     Evaluate                round((${Chambr_Analog} * 820) / abs(${PullUpValueAN3} - ${Chambr_Analog})) + 1
    ...                     ELSE                    Set Variable            10000
    ${WatPmp_NTCRes}        Run Keyword If          abs(${PullUpValueAN4} - ${WatPmp_Analog}) != 0
    ...                     Evaluate                round((${WatPmp_Analog} * 820) / abs(${PullUpValueAN4} - ${WatPmp_Analog})) + 1
    ...                     ELSE                    Set Variable            10000
    # The following codes are calculating the temperature value of the NTCs
    ${MOSFET_NTCTmp}        Evaluate                round(1 / (1/298.15 + 1/3435 * math.log(${MOSFET_NTCRes} / 10000)) - 273.15, 1)
    ${SBCTLE_NTCTmp}        Evaluate                round(1 / (1/298.15 + 1/3950 * math.log(${SBCTLE_NTCRes} / 10000)) - 273.15, 1)
    ${Chambr_NTCTmp}        Evaluate                round(1 / (1/298.15 + 1/3435 * math.log(${Chambr_NTCRes} / 10000)) - 273.15, 1)
    ${WatPmp_NTCTmp}        Evaluate                round(1 / (1/298.15 + 1/3435 * math.log(${WatPmp_NTCRes} / 10000)) - 273.15, 1)
    Set List Value          ${TMP}                  0                       ${MOSFET_NTCTmp}
    Set List Value          ${TMP}                  1                       ${SBCTLE_NTCTmp}
    Set List Value          ${TMP}                  2                       ${Chambr_NTCTmp}
    Set List Value          ${TMP}                  3                       ${WatPmp_NTCTmp}

Error_Mng
    [Arguments]             ${cmd_indx}             ${cmd_stat}
    run keyword if          ${cmd_indx} == ${CMD_INDICES.Excel_Crt}    fail                    ExcelCreationEror!!! ${cmd_stat[1]}.
    ...                     ELSE IF                 ${cmd_indx} == ${CMD_INDICES.Excel_Wrt}    fail                    ExcelWriteEror!!! ${cmd_stat[1]}.
    ...                     ELSE IF                 ${cmd_indx} == ${CMD_INDICES.CMD_TogglePort_A5}    fail                    Test failed in TogglePort_A5 command: ${cmd_stat}.
    ...                     ELSE IF                 ${cmd_indx} == ${CMD_INDICES.CMD_PWMoutput_A5}    fail                    Test failed in PWMoutput_A5 command: ${cmd_stat}.
    ...                     ELSE IF                 ${cmd_indx} == ${CMD_INDICES.CMD_work_mode}    fail                    Test failed in CMD_work_mode command: ${cmd_stat}.
    ...                     ELSE IF                 ${cmd_indx} == ${CMD_INDICES.EXP_DinGet}    fail                    Test failed in EXP_DinGet command: ${cmd_stat}
    ...                     ELSE IF                 ${cmd_indx} == ${CMD_INDICES.EXP_DoSet}    fail                    Test failed in EXP_DoSet command: ${cmd_stat}
    ...                     ELSE IF                 ${cmd_indx} == ${CMD_INDICES.EXP_FoSet}    fail                    Test failed in EXP_FoSet command: ${cmd_stat}
    ...                     ELSE IF                 ${cmd_indx} == ${CMD_INDICES.EXP_AoSet}    fail                    Test failed in EXP_AoSet command: ${cmd_stat}
    ...                     ELSE IF                 ${cmd_indx} == ${CMD_INDICES.CMD_ReportRequest_TLE8888_Diag}    fail                    Test failed in reportRequest_TLE8888 command: ${cmd_stat}
    ...                     ELSE IF                 ${cmd_indx} == ${CMD_INDICES.CMD_ReportRequest_ETC_Diag}    fail                    Test failed in reportRequest_ETC command: ${cmd_stat}
    ...                     ELSE                    log to console          unknownERROR

CalculateCheckSum
    [Arguments]             ${inputs_list}
    ${inputs_count} =       Get Length              ${inputs_list}
    ${checksum} =           Set Variable            ${inputs_list}[0]
    FOR                     ${index}                IN RANGE                1                       ${inputs_count}
        ${checksum}             Evaluate                ${checksum} ^ ${inputs_list}[${index}]
    END
    Append To List          ${inputs_list}          ${checksum}
    [Return]                ${inputs_list}

RegWrite_TLE
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-64%3A+TLE8888+Register+Write
    [Arguments]             ${register}             ${value}
    ${command}              Copy List               ${CMD_RegWrite_TLE8888}
    Append To List          ${command}              ${register}             ${value}
    # ${checksum}           CalculateCheckSum       ${command}
    # Append To List        ${command}              ${checksum}
    ${command}              SendGiveDatafromECU     ${command}
    Run Keyword If          ${command} == ${False}    Log To Console          TogglePort was failed    ELSE                    Log To Console         TogglePort was successful

Start Work mode
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-053%3A+Mode+Select
    [Arguments]             ${work_mode}            ${clear}                ${autostart}
    ${command}              Copy List               ${CMD_work_mode}
    Append To List          ${command}              ${work_mode}            ${clear}                ${autostart}
    Log To Console          ${command}
    ${command}              SendGiveDatafromECU     ${command}
    Run Keyword If          ${command} == ${False}    Log To Console          set work mode to ${work_mode} was failed    ELSE                    Log To Console         set work mode to ${work_mode} was successful

TogglePort
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-006%3A+Toggle+Port
    [Arguments]             ${GPIO_ID}              ${HighTimeMilisec}      ${LowTimeMilisec}
    Return From Keyword If    ${GPIO_ID} == ${-1}
    ${MSB}                  Evaluate                ${GPIO_ID} // ${256}
    ${LSB}                  Evaluate                ${GPIO_ID} % ${256}
    ${command}              Copy List               ${CMD_TogglePort}
    Append To List          ${command}              ${MSB}                  ${LSB}                  ${HighTimeMilisec}      ${LowTimeMilisec}
    ${command}              SendGiveDatafromECU     ${command}
    Run Keyword If          ${command} == ${False}  Log To Console          TogglePort was failed    ELSE                    Log To Console         TogglePort was successful

PWMoutput
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-032%3A+PWM+Output
    [Arguments]             ${CHANNEL_ID}           ${HighTimeMicrosec}     ${LowTimeMicrosec}
    Return From Keyword If    ${CHANNEL_ID} == ${-1}
    Return From Keyword If    ${HighTimeMicrosec} > ${65536} or ${HighTimeMicrosec} < ${1}
    Return From Keyword If    ${LowTimeMicrosec} > ${65536}    or ${LowTimeMicrosec} < ${1}
    ${H_MSB}                Evaluate                ${HighTimeMicrosec} // ${256}
    ${H_LSB}                Evaluate                ${HighTimeMicrosec} % ${256}
    ${L_MSB}                Evaluate                ${LowTimeMicrosec} // ${256}
    ${L_LSB}                Evaluate                ${LowTimeMicrosec} % ${256}
    ${command}              Copy List               ${CMD_PWMoutput}
    Append To List          ${command}              ${CHANNEL_ID}    ${H_MSB}    ${H_LSB}    ${L_MSB}    ${L_LSB}
    ${command}              SendGiveDatafromECU     ${command}
    Run Keyword If          ${command} == ${False}    Log To Console          PWMoutput was failed    ELSE    Log To Console    PWMoutput was successful

ReadAllAnalogInputs
    ${command}              Copy List               ${CMD_ReadAllAnalogInputs}
    ${ret}                  SendGiveDatafromECU     ${command}
    Run Keyword If          ${ret} == ${False}      Run Keywords            Log To Console          ReadAllAnalogInputs was failed    AND    Return From Keyword
    ${AnalogInputs_A5.V_IGK}              Evaluate                ${ret}[1][4]  + ${ret}[1][5]  * ${256}
    ${AnalogInputs_A5.V_EL}               Evaluate                ${ret}[1][6]  + ${ret}[1][7]  * ${256}
    ${AnalogInputs_A5.BAP_OUT}            Evaluate                ${ret}[1][8]  + ${ret}[1][9]  * ${256}
    ${AnalogInputs_A5.IGC_DIAG}           Evaluate                ${ret}[1][10] + ${ret}[1][11] * ${256}
    ${AnalogInputs_A5.FAN_DIAG}           Evaluate                ${ret}[1][12] + ${ret}[1][13] * ${256}
    ${AnalogInputs_A5.FTL}                Evaluate                ${ret}[1][14] + ${ret}[1][15] * ${256}
    ${AnalogInputs_A5.TCO}                Evaluate                ${ret}[1][16] + ${ret}[1][17] * ${256}
    ${AnalogInputs_A5.MAP}                Evaluate                ${ret}[1][18] + ${ret}[1][19] * ${256}
    ${AnalogInputs_A5.TIA}                Evaluate                ${ret}[1][20] + ${ret}[1][21] * ${256}
    ${AnalogInputs_A5.VLS_DOWN}           Evaluate                ${ret}[1][22] + ${ret}[1][23] * ${256}
    ${AnalogInputs_A5.VLS_UP}             Evaluate                ${ret}[1][24] + ${ret}[1][25] * ${256}
    ${AnalogInputs_A5.CRUISE_CTL}         Evaluate                ${ret}[1][26] + ${ret}[1][27] * ${256}
    ${AnalogInputs_A5.PVS1}               Evaluate                ${ret}[1][27] + ${ret}[1][28] * ${256}
    ${AnalogInputs_A5.PVS2}               Evaluate                ${ret}[1][29] + ${ret}[1][30] * ${256}
    ${AnalogInputs_A5.TPS1}               Evaluate                ${ret}[1][31] + ${ret}[1][32] * ${256}
    ${AnalogInputs_A5.TPS2}               Evaluate                ${ret}[1][33] + ${ret}[1][34] * ${256}
    ${AnalogInputs_A5.PUT}                Evaluate                ${ret}[1][35] + ${ret}[1][36] * ${256}
    ${AnalogInputs_A5.BRAKE_VACCUM}       Evaluate                ${ret}[1][37] + ${ret}[1][38] * ${256}
    ${AnalogInputs_A5.MODE_SW}            Evaluate                ${ret}[1][39] + ${ret}[1][40] * ${256}
    ${AnalogInputs_A5.ST_REQ}             Evaluate                ${ret}[1][41] + ${ret}[1][42] * ${256}
    ${AnalogInputs_A5.CRK_DIAG}           Evaluate                ${ret}[1][43] + ${ret}[1][44] * ${256}
    ${AnalogInputs_A5.NEUTRAL_GEAR_SW}    Evaluate                ${ret}[1][45] + ${ret}[1][46] * ${256}
    ${AnalogInputs_A5.PRS_PSTE}           Evaluate                ${ret}[1][47] + ${ret}[1][48] * ${256}
    ${AnalogInputs_A5.KNK}                Evaluate                ${ret}[1][49] + ${ret}[1][50] * ${256}
    log to console          ANALOGS -> V_IGK=${AnalogInputs_A5.V_IGK}, V_EL=${AnalogInputs_A5.V_EL}, BAP_OUT=${AnalogInputs_A5.BAP_OUT}, IGC_DIAG=${AnalogInputs_A5.IGC_DIAG},
    log to console          FAN_DIAG=${AnalogInputs_A5.FAN_DIAG}, FTL=${AnalogInputs_A5.FTL}, TCO=${AnalogInputs_A5.TCO}, MAP=${AnalogInputs_A5.MAP}, TIA=${AnalogInputs_A5.TIA},
    log to console          VLS_DOWN=${AnalogInputs_A5.VLS_DOWN}, VLS_UP=${AnalogInputs_A5.VLS_UP}, CRUISE_CTL=${AnalogInputs_A5.CRUISE_CTL}, PVS1=${AnalogInputs_A5.PVS1},
    log to console          PVS2=${AnalogInputs_A5.PVS2}, TPS1=${AnalogInputs_A5.TPS1}, TPS2=${AnalogInputs_A5.TPS2}, PUT=${AnalogInputs_A5.PUT}, BRAKE_VACCUM=${AnalogInputs_A5.BRAKE_VACCUM},
    log to console          MODE_SW=${AnalogInputs_A5.MODE_SW}, ST_REQ=${AnalogInputs_A5.ST_REQ}, CRK_DIAG=${AnalogInputs_A5.CRK_DIAG}, NEUTRAL_GEAR_SW=${AnalogInputs_A5.NEUTRAL_GEAR_SW},
    log to console          PRS_PSTE=${AnalogInputs_A5.PRS_PSTE}, KNK=${AnalogInputs_A5.KNK}

ReportRequest_AnalogInputs
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-055%3A+Report+Request
    ...                     link: http://192.168.5.62:8090/display/ECMT2/0x10+-+Analog+Inputs+Report+Response
    [Arguments]             ${Clear}
    Return From Keyword If    ${Clear} > ${1}         or ${Clear} < ${0}
    ${command}              Copy List               ${CMD_ReportRequest_AnalogInputs}
    Append To List          ${command}              ${Clear}
    ${ret}                  SendGiveDatafromECU     ${command}
    Run Keyword If          ${ret} == ${False}      Run Keywords            Log To Console          ReportRequest_AnalogInputs was failed    AND    Return From Keyword
    ${AnalogInputs_A5.V_IGK_MIN}              Evaluate                round((${ret}[1][4] * ${256} + ${ret}[1][5]) / 1024 * 5 * ${V_IGK_coef}, 3)
    ${AnalogInputs_A5.V_IGK_MAX}              Evaluate                round((${ret}[1][6] * ${256} + ${ret}[1][7]) / 1024 * 5 * ${V_IGK_coef}, 3)
    ${AnalogInputs_A5.V_EL_MIN}               Evaluate                round((${ret}[1][8] * ${256} + ${ret}[1][9]) / 1024 * 5 * ${V_EL_coef}, 3)
    ${AnalogInputs_A5.V_EL_MAX}               Evaluate                round((${ret}[1][10] * ${256} + ${ret}[1][11]) / 1024 * 5 * ${V_EL_coef}, 3)
    ${AnalogInputs_A5.BAP_OUT_MIN}            Evaluate                round(${ret}[1][12] * ${256} + ${ret}[1][13], 3)
    ${AnalogInputs_A5.BAP_OUT_MAX}            Evaluate                round(${ret}[1][14] * ${256} + ${ret}[1][15], 3)
    ${AnalogInputs_A5.IGC_DIAG_MIN}           Evaluate                round((${ret}[1][16] * ${256} + ${ret}[1][17]) / 1024 * 5, 3)
    ${AnalogInputs_A5.IGC_DIAG_MAX}           Evaluate                round((${ret}[1][18] * ${256} + ${ret}[1][19]) / 1024 * 5, 3)
    ${AnalogInputs_A5.FAN_DIAG_MIN}           Evaluate                round((${ret}[1][20] * ${256} + ${ret}[1][21]) / 1024 * 5, 3)
    ${AnalogInputs_A5.FAN_DIAG_MAX}           Evaluate                round((${ret}[1][22] * ${256} + ${ret}[1][23]) / 1024 * 5, 3)
    ${AnalogInputs_A5.FTL_MIN}                Evaluate                round((${ret}[1][24] * ${256} + ${ret}[1][25]) / 1024 * 5, 3)
    ${AnalogInputs_A5.FTL_MAX}                Evaluate                round((${ret}[1][26] * ${256} + ${ret}[1][27]) / 1024 * 5, 3)
    ${AnalogInputs_A5.TCO_MIN}                Evaluate                round((${ret}[1][28] * ${256} + ${ret}[1][29]) / 1024 * 5, 3)
    ${AnalogInputs_A5.TCO_MAX}                Evaluate                round((${ret}[1][30] * ${256} + ${ret}[1][31]) / 1024 * 5, 3)
    ${AnalogInputs_A5.MAP_MIN}                Evaluate                round((${ret}[1][32] * ${256} + ${ret}[1][33]) / 1024 * 5, 3)
    ${AnalogInputs_A5.MAP_MAX}                Evaluate                round((${ret}[1][34] * ${256} + ${ret}[1][35]) / 1024 * 5, 3)
    ${AnalogInputs_A5.TIA_MIN}                Evaluate                round((${ret}[1][36] * ${256} + ${ret}[1][37]) / 1024 * 5, 3)
    ${AnalogInputs_A5.TIA_MAX}                Evaluate                round((${ret}[1][38] * ${256} + ${ret}[1][39]) / 1024 * 5, 3)
    ${AnalogInputs_A5.VLS_DOWN_MIN}           Evaluate                round((${ret}[1][40] * ${256} + ${ret}[1][41]) / 1024 * 5, 3)
    ${AnalogInputs_A5.VLS_DOWN_MAX}           Evaluate                round((${ret}[1][42] * ${256} + ${ret}[1][43]) / 1024 * 5, 3)
    ${AnalogInputs_A5.VLS_UP_MIN}             Evaluate                round((${ret}[1][44] * ${256} + ${ret}[1][45]) / 1024 * 5, 3)
    ${AnalogInputs_A5.VLS_UP_MAX}             Evaluate                round((${ret}[1][46] * ${256} + ${ret}[1][47]) / 1024 * 5, 3)
    ${AnalogInputs_A5.CRUISE_CTL_MIN}         Evaluate                round((${ret}[1][48] * ${256} + ${ret}[1][49]) / 1024 * 5, 3)
    ${AnalogInputs_A5.CRUISE_CTL_MAX}         Evaluate                round((${ret}[1][50] * ${256} + ${ret}[1][51]) / 1024 * 5, 3)
    ${AnalogInputs_A5.PVS1_MIN}               Evaluate                round((${ret}[1][52] * ${256} + ${ret}[1][53]) / 1024 * 5, 3)
    ${AnalogInputs_A5.PVS1_MAX}               Evaluate                round((${ret}[1][54] * ${256} + ${ret}[1][55]) / 1024 * 5, 3)
    ${AnalogInputs_A5.PVS2_MIN}               Evaluate                round((${ret}[1][56] * ${256} + ${ret}[1][57]) / 1024 * 5, 3)
    ${AnalogInputs_A5.PVS2_MAX}               Evaluate                round((${ret}[1][58] * ${256} + ${ret}[1][59]) / 1024 * 5, 3)
    ${AnalogInputs_A5.TPS1_MIN}               Evaluate                round((${ret}[1][60] * ${256} + ${ret}[1][61]) / 1024 * 5, 3)
    ${AnalogInputs_A5.TPS1_MAX}               Evaluate                round((${ret}[1][62] * ${256} + ${ret}[1][63]) / 1024 * 5, 3)
    ${AnalogInputs_A5.TPS2_MIN}               Evaluate                round((${ret}[1][64] * ${256} + ${ret}[1][65]) / 1024 * 5, 3)
    ${AnalogInputs_A5.TPS2_MAX}               Evaluate                round((${ret}[1][66] * ${256} + ${ret}[1][67]) / 1024 * 5, 3)
    ${AnalogInputs_A5.PUT_MIN}                Evaluate                round((${ret}[1][68] * ${256} + ${ret}[1][69]) / 1024 * 5, 3)
    ${AnalogInputs_A5.PUT_MAX}                Evaluate                round((${ret}[1][70] * ${256} + ${ret}[1][71]) / 1024 * 5, 3)
    ${AnalogInputs_A5.BRAKE_VACCUM_MIN}       Evaluate                round((${ret}[1][72] * ${256} + ${ret}[1][73]) / 1024 * 5, 3)
    ${AnalogInputs_A5.BRAKE_VACCUM_MAX}       Evaluate                round((${ret}[1][74] * ${256} + ${ret}[1][75]) / 1024 * 5, 3)
    ${AnalogInputs_A5.MODE_SW_MIN}            Evaluate                round((${ret}[1][76] * ${256} + ${ret}[1][77]) / 1024 * 5, 3)
    ${AnalogInputs_A5.MODE_SW_MAX}            Evaluate                round((${ret}[1][78] * ${256} + ${ret}[1][79]) / 1024 * 5, 3)
    ${AnalogInputs_A5.ST_REQ_MIN}             Evaluate                round((${ret}[1][80] * ${256} + ${ret}[1][81]) / 1024 * 5, 3)
    ${AnalogInputs_A5.ST_REQ_MAX}             Evaluate                round((${ret}[1][82] * ${256} + ${ret}[1][83]) / 1024 * 5, 3)
    ${AnalogInputs_A5.CRK_DIAG_MIN}           Evaluate                round((${ret}[1][84] * ${256} + ${ret}[1][85]) / 1024 * 5, 3)
    ${AnalogInputs_A5.CRK_DIAG_MAX}           Evaluate                round((${ret}[1][86] * ${256} + ${ret}[1][87]) / 1024 * 5, 3)
    ${AnalogInputs_A5.NEUTRAL_GEAR_SW_MIN}    Evaluate                round((${ret}[1][88] * ${256} + ${ret}[1][89]) / 1024 * 5, 3)
    ${AnalogInputs_A5.NEUTRAL_GEAR_SW_MAX}    Evaluate                round((${ret}[1][90] * ${256} + ${ret}[1][91]) / 1024 * 5, 3)
    ${AnalogInputs_A5.PRS_PSTE_MIN}           Evaluate                round((${ret}[1][92] * ${256} + ${ret}[1][93]) / 1024 * 5, 3)
    ${AnalogInputs_A5.PRS_PSTE_MAX}           Evaluate                round((${ret}[1][94] * ${256} + ${ret}[1][95]) / 1024 * 5, 3)
    ${AnalogInputs_A5.KNK_MIN}                Evaluate                round((${ret}[1][96] * ${256} + ${ret}[1][97]) / 1024 * 5, 3)
    ${AnalogInputs_A5.KNK_MAX}                Evaluate                round((${ret}[1][98] * ${256} + ${ret}[1][99]) / 1024 * 5, 3)

ReSet_DigitalInputs
    # ACIN
    ${Ret_List}             Do Set                  ${3}                    ${0}
    # ACCPRS1
    ${Ret_List}             Do Set                  ${1}                    ${0}
    # ACCPRS2
    ${Ret_List}             Do Set                  ${2}                    ${0}
    # IGK
    # ${Ret_List}           Do Set                  ${0}                    ${0}
    # CLUSWI
    ${Ret_List}             Do Set                  ${8}                    ${0}
    # BLS
    ${Ret_List}             Do Set                  ${7}                    ${0}
    # BTS
    ${Ret_List}             Do Set                  ${6}                    ${0}
    # CLUSWI1
    ${Ret_List}             Do Set                  ${15}                   ${0}
    Set List Value          ${DigitalInput}         0                       ${0}
    Set List Value          ${DigitalInput}         1                       ${0}
    Set List Value          ${DigitalInput}         2                       ${0}
    Set List Value          ${DigitalInput}         3                       ${0}
    Set List Value          ${DigitalInput}         4                       ${0}
    Set List Value          ${DigitalInput}         5                       ${0}
    Set List Value          ${DigitalInput}         6                       ${0}
    Set List Value          ${DigitalInput}         7                       ${0}
    Set List Value          ${DigitalInput}         8                       ${0}
    
Set_DigitalInputs
    # ACIN
    ${Ret_List}             Do Set                  ${3}                    ${1}
    # ACCPRS1
    ${Ret_List}             Do Set                  ${1}                    ${1}
    # ACCPRS2
    ${Ret_List}             Do Set                  ${2}                    ${1}
    # IGK
    # ${Ret_List}           Do Set                  ${0}                    ${1}
    # CLUSWI
    ${Ret_List}             Do Set                  ${8}                    ${1}
    # BLS
    ${Ret_List}             Do Set                  ${7}                    ${1}
    # BTS
    ${Ret_List}             Do Set                  ${6}                    ${1}
    # CLUSWI1
    ${Ret_List}             Do Set                  ${15}                   ${1}
    Set List Value          ${DigitalInput}         0                       ${1}
    Set List Value          ${DigitalInput}         1                       ${1}
    Set List Value          ${DigitalInput}         2                       ${1}
    Set List Value          ${DigitalInput}         3                       ${1}
    Set List Value          ${DigitalInput}         4                       ${1}
    Set List Value          ${DigitalInput}         5                       ${1}
    Set List Value          ${DigitalInput}         6                       ${1}
    Set List Value          ${DigitalInput}         7                       ${1}
    Set List Value          ${DigitalInput}         8                       ${1}
    
AoSel_AnalogInputs
    [Arguments]             ${loopCounter}
    ${argus}                Evaluate                ${loopCounter} % 8
    FOR                     ${i}                    IN RANGE                ${0}                    ${16}
        Run Keyword If          ${argus} == 0           Ao Sel                  ${i}                    ${0}                    ${0}                   ${0}
        Run Keyword If          ${argus} == 1           Ao Sel                  ${i}                    ${1}                    ${0}                   ${0}
        Run Keyword If          ${argus} == 2           Ao Sel                  ${i}                    ${0}                    ${1}                   ${0}
        Run Keyword If          ${argus} == 3           Ao Sel                  ${i}                    ${0}                    ${0}                   ${1}
        Run Keyword If          ${argus} == 4           Ao Sel                  ${i}                    ${1}                    ${1}                   ${0}
        Run Keyword If          ${argus} == 5           Ao Sel                  ${i}                    ${1}                    ${0}                   ${1}
        Run Keyword If          ${argus} == 6           Ao Sel                  ${i}                    ${0}                    ${1}                   ${1}
        Run Keyword If          ${argus} == 7           Ao Sel                  ${i}                    ${1}                    ${1}                   ${1}
    END

ReportRequest_DigitalInputs
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-055%3A+Report+Request
    ...                     link: http://192.168.5.62:8090/display/ECMT2/0x20+-+Logic+Inputs+Report+Request
    [Arguments]             ${Clear}
    Return From Keyword If    ${Clear} > ${1}         or ${Clear} < ${0}
    ${command}              Copy List               ${CMD_ReportRequest_DigitalInputs}
    Append To List          ${command}              ${Clear}
    ${ret}                  SendGiveDatafromECU     ${command}
    Run Keyword If          ${ret} == ${False}      Run Keywords            Log To Console          ReportRequest_DigitalInputs was failed    AND    Return From Keyword
    Log To Console    ${ret}
    ${FirstByte}            Convert To Binary       ${ret}[1][4]            length=8
    ${SecondByte}           Convert To Binary       ${ret}[1][5]            length=8
    ${ThirdByte}           Convert To Binary       ${ret}[1][6]            length=8
    ${DigitalInputs_A5.ACIN_CurLev}                 Set Variable            ${FirstByte}[7]
    ${DigitalInputs_A5.ACIN_LevChang}               Set Variable            ${FirstByte}[6]
    ${DigitalInputs_A5.ACC_PRS1_CurLev}             Set Variable            ${FirstByte}[5]
    ${DigitalInputs_A5.ACC_PRS1_LevChang}           Set Variable            ${FirstByte}[4]
    ${DigitalInputs_A5.ACC_PRS2_CurLev}             Set Variable            ${FirstByte}[3]
    ${DigitalInputs_A5.ACC_PRS2_LevChang}           Set Variable            ${FirstByte}[2]
    ${DigitalInputs_A5.IGK_STATUS_CurLev}           Set Variable            ${FirstByte}[1]
    ${DigitalInputs_A5.IGK_STATUS_LevChang}         Set Variable            ${FirstByte}[0]
    
    ${DigitalInputs_A5.CLUTCH_SW_CurLev}            Set Variable            ${SecondByte}[7]
    ${DigitalInputs_A5.CLUTCH_SW_LevChang}          Set Variable            ${SecondByte}[6]
    ${DigitalInputs_A5.BRAKE_LIGHT_SW_CurLev}       Set Variable            ${SecondByte}[5]
    ${DigitalInputs_A5.BRAKE_LIGHT_SW_LevChang}     Set Variable            ${SecondByte}[4]
    ${DigitalInputs_A5.BRAKE_TEST_SW_CurLev}        Set Variable            ${SecondByte}[3]
    ${DigitalInputs_A5.BRAKE_TEST_SW_LevChang}      Set Variable            ${SecondByte}[2]
    ${DigitalInputs_A5.CLUTCH_SW2_CurLev}           Set Variable            ${SecondByte}[1]
    ${DigitalInputs_A5.CLUTCH_SW2_LevChang}         Set Variable            ${SecondByte}[0]
    
    ${DigitalInputs_A5.KICK_DOWN_SW_CurLev}         Set Variable            ${SecondByte}[7]
    ${DigitalInputs_A5.KICK_DOWN_SW_LevChang}       Set Variable            ${SecondByte}[6]
    [Return]                ${DigitalInputs_A5.ACIN_CurLev}    ${DigitalInputs_A5.ACIN_LevChang}
    ...                     ${DigitalInputs_A5.ACC_PRS1_CurLev}    ${DigitalInputs_A5.ACC_PRS1_LevChang}
    ...                     ${DigitalInputs_A5.ACC_PRS2_CurLev}    ${DigitalInputs_A5.ACC_PRS2_LevChang}
    ...                     ${DigitalInputs_A5.IGK_STATUS_CurLev}  ${DigitalInputs_A5.IGK_STATUS_LevChang}
    ...                     ${DigitalInputs_A5.CLUTCH_SW_CurLev}   ${DigitalInputs_A5.CLUTCH_SW_LevChang}
    ...                     ${DigitalInputs_A5.BRAKE_LIGHT_SW_CurLev}    ${DigitalInputs_A5.BRAKE_LIGHT_SW_LevChang}
    ...                     ${DigitalInputs_A5.BRAKE_TEST_SW_CurLev}     ${DigitalInputs_A5.BRAKE_TEST_SW_LevChang}
    ...                     ${DigitalInputs_A5.CLUTCH_SW2_CurLev}        ${DigitalInputs_A5.CLUTCH_SW2_LevChang}
    ...                     ${DigitalInputs_A5.KICK_DOWN_SW_CurLev}      ${DigitalInputs_A5.KICK_DOWN_SW_LevChang}

ReportRequest_FrequencyInputs
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-055%3A+Report+Request
    ...                     link: http://192.168.5.62:8090/display/ECMT2/0x30+-+Frequency+Inputs+Report+Request
    [Arguments]             ${Clear}
    Return From Keyword If  ${Clear} > ${1}         or ${Clear} < ${0}
    ${command}              Copy List               ${CMD_ReportRequest_FrequencyInputs}
    Append To List          ${command}              ${Clear}
    ${ret}                  SendGiveDatafromECU     ${command}
    Run Keyword If          ${ret} == ${False}      Run Keywords            Log To Console          ReportRequest_FrequencyInputs was failed   AND    Return From Keyword
    ${FrequencyInputs_A5.CAM_HIGH_MIN}        Evaluate                (${ret}[1][4] * ${256} + ${ret}[1][5]) * 4
    ${FrequencyInputs_A5.CAM_HIGH_MAX}        Evaluate                (${ret}[1][6] * ${256} + ${ret}[1][7]) * 4
    ${FrequencyInputs_A5.CAM_LOW_MIN}         Evaluate                (${ret}[1][8] * ${256} + ${ret}[1][9]) * 4
    ${FrequencyInputs_A5.CAM_LOW_MAX}         Evaluate                (${ret}[1][10] * ${256} + ${ret}[1][11]) * 4
    ${FrequencyInputs_A5.WSS_HIGH_MIN}        Evaluate                (${ret}[1][12] * ${256} + ${ret}[1][13]) * 1
    ${FrequencyInputs_A5.WSS_HIGH_MAX}        Evaluate                (${ret}[1][14] * ${256} + ${ret}[1][15]) * 1
    ${FrequencyInputs_A5.WSS_LOW_MIN}         Evaluate                (${ret}[1][16] * ${256} + ${ret}[1][17]) * 1
    ${FrequencyInputs_A5.WSS_LOW_MAX}         Evaluate                (${ret}[1][18] * ${256} + ${ret}[1][19]) * 1
    ${FrequencyInputs_A5.ALT_MON_HIGH_MIN}    Evaluate                (${ret}[1][20] * ${256} + ${ret}[1][21]) * 1
    ${FrequencyInputs_A5.ALT_MON_HIGH_MAX}    Evaluate                (${ret}[1][22] * ${256} + ${ret}[1][23]) * 1
    ${FrequencyInputs_A5.ALT_MON_LOW_MIN}     Evaluate                (${ret}[1][24] * ${256} + ${ret}[1][25]) * 1
    ${FrequencyInputs_A5.ALT_MON_LOW_MAX}     Evaluate                (${ret}[1][26] * ${256} + ${ret}[1][27]) * 1
    ${FrequencyInputs_A5.CAM_EX_HIGH_MIN}     Evaluate                (${ret}[1][28] * ${256} + ${ret}[1][29]) * 1
    ${FrequencyInputs_A5.CAM_EX_HIGH_MAX}     Evaluate                (${ret}[1][30] * ${256} + ${ret}[1][31]) * 1
    ${FrequencyInputs_A5.CAM_EX_LOW_MIN}      Evaluate                (${ret}[1][32] * ${256} + ${ret}[1][33]) * 1
    ${FrequencyInputs_A5.CAM_EX_LOW_MAX}      Evaluate                (${ret}[1][34] * ${256} + ${ret}[1][35]) * 1
    ${FrequencyInputs_A5.CRANK_HIGH_MIN}      Evaluate                (${ret}[1][36] * ${256} + ${ret}[1][37]) * 1
    ${FrequencyInputs_A5.CRANK_HIGH_MAX}      Evaluate                (${ret}[1][38] * ${256} + ${ret}[1][39]) * 1
    ${FrequencyInputs_A5.CRANK_LOW_MIN}       Evaluate                (${ret}[1][40] * ${256} + ${ret}[1][41]) * 1
    ${FrequencyInputs_A5.CRANK_LOW_MAX}       Evaluate                (${ret}[1][42] * ${256} + ${ret}[1][43]) * 1
    [Return]
    ...                     ${FrequencyInputs_A5.CAM_HIGH_MIN}
    ...                     ${FrequencyInputs_A5.CAM_HIGH_MAX}
    ...                     ${FrequencyInputs_A5.CAM_LOW_MIN}
    ...                     ${FrequencyInputs_A5.CAM_LOW_MAX}
    ...                     ${FrequencyInputs_A5.WSS_HIGH_MIN}
    ...                     ${FrequencyInputs_A5.WSS_HIGH_MAX}
    ...                     ${FrequencyInputs_A5.WSS_LOW_MIN}
    ...                     ${FrequencyInputs_A5.WSS_LOW_MAX}
    ...                     ${FrequencyInputs_A5.ALT_MON_HIGH_MIN}
    ...                     ${FrequencyInputs_A5.ALT_MON_HIGH_MAX}
    ...                     ${FrequencyInputs_A5.ALT_MON_LOW_MIN}
    ...                     ${FrequencyInputs_A5.ALT_MON_LOW_MAX}
    ...                     ${FrequencyInputs_A5.CAM_EX_HIGH_MIN}
    ...                     ${FrequencyInputs_A5.CAM_EX_HIGH_MAX}
    ...                     ${FrequencyInputs_A5.CAM_EX_LOW_MIN}
    ...                     ${FrequencyInputs_A5.CAM_EX_LOW_MAX}
    ...                     ${FrequencyInputs_A5.CRANK_HIGH_MIN}
    ...                     ${FrequencyInputs_A5.CRANK_HIGH_MAX}
    ...                     ${FrequencyInputs_A5.CRANK_LOW_MIN}
    ...                     ${FrequencyInputs_A5.CRANK_LOW_MAX}

ReportRequest_TLE8888_Diag
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-055%3A+Report+Request
    ...                     link: http://192.168.5.62:8090/display/ECMT2/0x53+-+SBC+TLE8888+Report+Request
    [Arguments]             ${Clear}
    Return From Keyword If    ${Clear} > ${1}         or ${Clear} < ${0}
    ${command}              Copy List               ${CMD_ReportRequest_TLE8888_Diag}
    Append To List          ${command}              ${Clear}
    ${ret}                  SendGiveDatafromECU     ${command}
    Run Keyword If          ${ret} == ${False}      Run Keywords            Log To Console          ReportRequest_TLE8888_Diag was failed   AND    Return From Keyword
    ${TLE8888_Diag_A5.Byte0}    Convert To Binary       ${ret}[1][4]            length=8
    ${TLE8888_Diag_A5.Byte1}    Convert To Binary       ${ret}[1][5]            length=8
    ${TLE8888_Diag_A5.Byte2}    Convert To Binary       ${ret}[1][6]            length=8
    ${TLE8888_Diag_A5.Byte3}    Convert To Binary       ${ret}[1][7]            length=8
    ${TLE8888_Diag_A5.Byte4}    Convert To Binary       ${ret}[1][8]            length=8
    ${TLE8888_Diag_A5.Byte5}    Convert To Binary       ${ret}[1][9]            length=8
    ${TLE8888_Diag_A5.Byte6}    Convert To Binary       ${ret}[1][10]           length=8
    ${TLE8888_Diag_A5.Byte7}    Convert To Binary       ${ret}[1][11]           length=8
    ${TLE8888_Diag_A5.Diag0}    Convert To Binary       ${ret}[1][20]           length=8
    # Log To Console    ${ret}
    # Log To Console    1:${ret}[1][1]
    # Log To Console    5:${ret}[1][5]
    # Log To Console    20:${ret}[1][20]
    # Log To Console    21:${ret}[1][21]
    # Log To Console    19:${ret}[1][19]
    # Log To Console    18:${ret}[1][18]
    # Log To Console    17:${ret}[1][17]
    # Log To Console    16:${ret}[1][16]
    
TranslateTLE8888Diag
    [Arguments]             ${MSb}                  ${LSb}
    ${dummy}                Run Keyword If          ${MSb}== 0 and ${LSb}== 0    Set Variable            ${DiagResponse.NoFailure}
    ...                     ELSE IF                 ${MSb}== 0 and ${LSb}== 1    Set Variable            ${DiagResponse.ShortToBAT}
    ...                     ELSE IF                 ${MSb}== 1 and ${LSb}== 0    Set Variable            ${DiagResponse.OpenLoad}
    ...                     ELSE IF                 ${MSb}== 1 and ${LSb}== 1    Set Variable            ${DiagResponse.ShortToGround}
    [Return]                ${dummy}

ReportRequest_ETC_Diag
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-055%3A+Report+Request
    ...                     link: http://192.168.5.62:8090/display/ECMT2/0x60+-+L9960+Report+Request
    [Arguments]             ${Clear}
    Return From Keyword If    ${Clear} > ${1}         or ${Clear} < ${0}
    ${command}              Copy List               ${CMD_ReportRequest_ETC_Diag}
    Append To List          ${command}              ${Clear}
    ${ret}                  SendGiveDatafromECU     ${command}
    Run Keyword If          ${ret} == ${False}      Run Keywords            Log To Console          ReportRequest_ETC_Diag was failed   AND    Return From Keyword
    ${ETC_Diag_A5.overcurrentMonitoring}    Evaluate                ${ret}[1][4] * 256 + ${ret}[1][5]
    ${ETC_Diag_A5.overcurrentMonitoring}    Convert To Binary       ${ETC_Diag_A5.overcurrentMonitoring}    length=16
    ${ETC_Diag_A5.statesResponse1}          Evaluate                ${ret}[1][16] * 256 + ${ret}[1][17]
    ${ETC_Diag_A5.statesResponse1}          Convert To Binary       ${ETC_Diag_A5.statesResponse1}    length=16
    ${ETC_Diag_A5.statesResponse2}          Evaluate                ${ret}[1][18] * 256 + ${ret}[1][19]
    ${ETC_Diag_A5.statesResponse2}          Convert To Binary       ${ETC_Diag_A5.statesResponse2}    length=16
    ${ETC_Diag_A5.statesResponse3}          Evaluate                ${ret}[1][20] * 256 + ${ret}[1][21]
    ${ETC_Diag_A5.statesResponse3}          Convert To Binary       ${ETC_Diag_A5.statesResponse3}    length=16
    ${ETC_Diag_A5.VDD_OV_UV}                Evaluate                ${ret}[1][22] * 256 + ${ret}[1][23]
    ${ETC_Diag_A5.VDD_OV_UV}                Convert To Binary       ${ETC_Diag_A5.VDD_OV_UV}    length=16

ReportRequest_KP254
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-055%3A+Report+Request
    ...                     link: http://192.168.5.62:8090/display/ECMT2/0x92+-+KP254+Report+Request
    [Arguments]                ${Clear}
    Return From Keyword If     ${Clear} > ${1}         or ${Clear} < ${0}
    ${command}                 Copy List               ${CMD_ReportRequest_KP254}
    Append To List             ${command}              ${Clear}
    ${ret}                     SendGiveDatafromECU     ${command}
    Run Keyword If             ${ret} == ${False}      Run Keywords            Log To Console          ReportRequest_KP254 was failed   AND    Return From Keyword
    ${KP254_A5.PRESSURE}       Evaluate                ${ret}[1][4] * 256 + ${ret}[1][5]
    ${KP254_A5.TEMPERATURE}    Evaluate                ${ret}[1][6] * 256 + ${ret}[1][7]
    ${KP254_A5.PRESSURE}       Convert To Binary       ${KP254_A5.PRESSURE}    length=16
    ${KP254_A5.TEMPERATURE}    Convert To Binary       ${KP254_A5.TEMPERATURE}    length=16
    ${KP254_A5.DIAG_H}         Convert To Binary       ${ret}[1][8]            length=8
    ${KP254_A5.DIAG_L}         Convert To Binary       ${ret}[1][9]            length=8
    [Return]                   ${KP254_A5.PRESSURE}    ${KP254_A5.TEMPERATURE}    ${KP254_A5.DIAG_H}      ${KP254_A5.DIAG_L}

ReportRequest_MCUTemperature
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/0x91+-+Temperature+Report+Request
    [Arguments]               ${Clear}
    Return From Keyword If    ${Clear} > ${1}         or ${Clear} < ${0}
    ${command}                Copy List               ${CMD_ReportRequest_MCUTemperature}
    Append To List            ${command}              ${Clear}
    ${ret}                    SendGiveDatafromECU     ${command}
    Run Keyword If            ${ret} == ${False}      Run Keywords            Log To Console          ReportRequest_KP254 was failed   AND    Return From Keyword
    ${MCUTemperature}         Evaluate                ${ret}[1][4] + ${0.1} * ${ret}[1][5]
    Set List Value            ${TMP}                  4                       ${MCUTemperature}
    # Log To Console          ${MCUTemperature}

ReportRequest_InternalStatus
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-055%3A+Report+Request
    ...                     link: http://192.168.5.62:8090/display/ECMT2/0x00+-+Internal+Status+Report+Response
    [Arguments]             ${Clear}
    Return From Keyword If    ${Clear} > ${1}         or ${Clear} < ${0}
    ${command}              Copy List               ${CMD_ReportRequest_InternalStatus}
    Append To List          ${command}              ${Clear}
    ${ret}                  SendGiveDatafromECU     ${command}
    Run Keyword If          ${ret} == ${False}   Run Keywords            Log To Console    ReportRequest_InternalStatus was failed    AND    Return From Keyword
    # Log To Console    ${ret}
    ${InternalStatus_A5.CurrentRunningMode}      Set Variable            ${ret}[1][4]
    ${InternalStatus_A5.MClastReset}             Set Variable            ${ret}[1][5]
    ${InternalStatus_A5.MUlastReset}             Set Variable            ${ret}[1][6]
    ${InternalStatus_A5.MUresetCount}            Set Variable            ${ret}[1][7]
    ${InternalStatus_A5.MUerrorCode}             Set Variable            ${ret}[1][8]
    ${InternalStatus_A5.MUstate}                 Set Variable            ${ret}[1][9]
    ${InternalStatus_A5.MCExRstCount}            Set Variable            ${ret}[1][10]

ReportRequest_FrequencyOutputs_1
    ${GIV1}                 Fin Get                 ${4}
    ${CP}                   Fin Get                 ${2}
    ${GIV3}                 Fin Get                 ${6}
    ${IGC3}                 Fin Get                 ${10}
    ${IGC4}                 Fin Get                 ${11}
    ${INJ2}                 Fin Get                 ${13}
    ${INJ3}                 Fin Get                 ${14}
    ${WG}                   Fin Get                 ${20}
    ${LSHDOWN}              Fin Get                 ${17}
    ${INJ1}                 Fin Get                 ${12}
    ${IGC1}                 Fin Get                 ${8}
    Set List Value          ${Freq_Outputs}         1                       ${CP}[0] ${CP}[1]
    Set List Value          ${Freq_Outputs}         3                       ${GIV1}[0] ${GIV1}[1]
    Set List Value          ${Freq_Outputs}         5                       ${GIV3}[0] ${GIV3}[1]
    Set List Value          ${Freq_Outputs}         7                       ${IGC1}[0] ${IGC1}[1]
    Set List Value          ${Freq_Outputs}         9                       ${IGC3}[0] ${IGC3}[1]
    Set List Value          ${Freq_Outputs}         10                      ${IGC4}[0] ${IGC4}[1]
    Set List Value          ${Freq_Outputs}         11                      ${INJ1}[0] ${INJ1}[1]
    Set List Value          ${Freq_Outputs}         12                      ${INJ2}[0] ${INJ2}[1]
    Set List Value          ${Freq_Outputs}         13                      ${INJ3}[0] ${INJ3}[1]
    Set List Value          ${Freq_Outputs}         16                      ${LSHDOWN}[0] ${LSHDOWN}[1]
    Set List Value          ${Freq_Outputs}         19                      ${WG}[0] ${WG}[1]
    sleep                   100ms

ReportRequest_FrequencyOutputs_2
    ${LSHUP}                Fin Get                 ${18}
    ${IVVT}                 Fin Get                 ${16}
    ${IGC2}                 Fin Get                 ${9}
    ${EVVT}                 Fin Get                 ${3}
    ${ALT_CMD}              Fin Get                 ${1}
    ${GIV4}                 Fin Get                 ${7}
    ${GIV2}                 Fin Get                 ${5}
    ${WG_NEG}               Fin Get                 ${19}
    ${INJ4}                 Fin Get                 ${12}
    Set List Value          ${Freq_Outputs}         17                      ${LSHUP}[0] ${LSHUP}[1]
    Set List Value          ${Freq_Outputs}         15                      ${IVVT}[0] ${IVVT}[1]
    Set List Value          ${Freq_Outputs}         8                       ${IGC2}[0] ${IGC2}[1]
    Set List Value          ${Freq_Outputs}         2                       ${EVVT}[0] ${EVVT}[1]
    Set List Value          ${Freq_Outputs}         0                       ${ALT_CMD}[0] ${ALT_CMD}[1]
    Set List Value          ${Freq_Outputs}         6                       ${GIV4}[0] ${GIV4}[1]
    Set List Value          ${Freq_Outputs}         4                       ${GIV2}[0] ${GIV2}[1]
    Set List Value          ${Freq_Outputs}         18                      ${WG_NEG}[0] ${WG_NEG}[1]
    Set List Value          ${Freq_Outputs}         14                      ${INJ4}[0] ${INJ4}[1]
    sleep                   100ms

FinSelMux_ReSet
    Fin Sel                 ${1}                    ${0}                    ${0}                    ${0}
    Fin Sel                 ${2}                    ${0}                    ${0}                    ${0}
    Fin Sel                 ${3}                    ${0}                    ${0}                    ${0}

FinSelMux_Set
    Fin Sel                 ${1}                    ${1}                    ${1}                    ${1}
    Fin Sel                 ${2}                    ${1}                    ${1}                    ${1}
    Fin Sel                 ${3}                    ${1}                    ${1}                    ${1}

ReportRequest_MonitoringUnit
    [Documentation]         link: http://192.168.5.62:8090/display/ECMT2/ID-055%3A+Report+Request
    ...                     link: http://192.168.5.62:8090/display/ECMT2/0x40+-+Monitoring+Unit+Report+Request
    ...                     link: http://192.168.5.62:8090/display/ECMT/Monitoring+Unit+Report+Response
    [Arguments]             ${Clear}
    Return From Keyword If    ${Clear} > ${1}         or ${Clear} < ${0}
    ${command}              Copy List               ${CMD_ReportRequest_MonitoringUnit}
    Append To List          ${command}              ${Clear}
    ${ret}                  SendGiveDatafromECU     ${command}
    Run Keyword If          ${ret} == ${False}      Run Keywords    Log To Console    ReportRequest_MonitoringUnit was failed    AND    Return From Keyword
    ${MU_Part01}            Evaluate                ${ret}[1][4] * 256 + ${ret}[1][5]
    ${MU_Part01}            Convert To Binary       ${MU_Part01}            length=16
    ${MU_Part02}            Evaluate                ${ret}[1][6] * 256 + ${ret}[1][7]
    ${MU_Part02}            Convert To Binary       ${MU_Part02}            length=16
    ${MU_Part03}            Evaluate                ${ret}[1][8] * 256 + ${ret}[1][9]
    ${MU_Part03}            Convert To Binary       ${MU_Part03}            length=16
    ${MU_Part04}            Evaluate                ${ret}[1][10] * 256 + ${ret}[1][11]
    ${MU_Part04}            Convert To Binary       ${MU_Part04}            length=16
    ${MU_Part05}            Evaluate                ${ret}[1][12] * 256 + ${ret}[1][13]
    ${MU_Part05}            Convert To Binary       ${MU_Part05}            length=16
    ${MU_Part06}            Evaluate                ${ret}[1][14] * 256 + ${ret}[1][15]
    ${MU_Part06}            Convert To Binary       ${MU_Part06}            length=16
    ${MU_Part07}            Evaluate                ${ret}[1][16] * 256 + ${ret}[1][17]
    ${MU_Part07}            Convert To Binary       ${MU_Part07}            length=16
    ${MU_Part08}            Evaluate                ${ret}[1][18] * 256 + ${ret}[1][19]
    ${MU_Part08}            Convert To Binary       ${MU_Part08}            length=16
    ${MU_Part09}            Evaluate                ${ret}[1][20] * 256 + ${ret}[1][21]
    ${MU_Part09}            Convert To Binary       ${MU_Part09}            length=16
    ${MU_Part10}            Evaluate                ${ret}[1][22] * 256 + ${ret}[1][23]
    ${MU_Part10}            Convert To Binary       ${MU_Part10}            length=16
    ${MU_Part11}            Evaluate                ${ret}[1][24] * 256 + ${ret}[1][25]
    ${MU_Part11}            Convert To Binary       ${MU_Part11}            length=16
    ${MU_Part12}            Evaluate                ${ret}[1][26] * 256 + ${ret}[1][27]
    ${MU_Part12}            Convert To Binary       ${MU_Part12}            length=16
    ${MU_Part13}            Evaluate                ${ret}[1][28] * 256 + ${ret}[1][29]
    ${MU_Part13}            Convert To Binary       ${MU_Part13}            length=16
    ${MU_Part14}            Evaluate                ${ret}[1][30] * 256 + ${ret}[1][31]
    ${MU_Part14}            Convert To Binary       ${MU_Part14}            length=16
    ${MU_Part15}            Evaluate                ${ret}[1][32] * 256 + ${ret}[1][33]
    ${MU_Part15}            Convert To Binary       ${MU_Part15}            length=16
    ${MU_Part16}            Evaluate                ${ret}[1][34] * 256 + ${ret}[1][35]
    ${MU_Part16}            Convert To Binary       ${MU_Part16}            length=16
    Set List Value          ${MU}                   0                       ${MU_Part01}
    Set List Value          ${MU}                   1                       ${MU_Part02}
    Set List Value          ${MU}                   2                       ${MU_Part03}
    Set List Value          ${MU}                   3                       ${MU_Part04}
    Set List Value          ${MU}                   4                       ${MU_Part05}
    Set List Value          ${MU}                   5                       ${MU_Part06}
    Set List Value          ${MU}                   6                       ${MU_Part07}
    Set List Value          ${MU}                   7                       ${MU_Part08}
    Set List Value          ${MU}                   8                       ${MU_Part09}
    Set List Value          ${MU}                   9                       ${MU_Part10}
    Set List Value          ${MU}                   10                      ${MU_Part11}
    Set List Value          ${MU}                   11                      ${MU_Part12}
    Set List Value          ${MU}                   12                      ${MU_Part13}
    Set List Value          ${MU}                   13                      ${MU_Part14}
    Set List Value          ${MU}                   14                      ${MU_Part15}
    Set List Value          ${MU}                   15                      ${MU_Part16}
    [Return]                ${MU}

TranslateETCovercurrentDiag
    [Arguments]             ${MSb}                  ${LSb}
    ${dummy}                Run Keyword If          ${MSb}== 0 and ${LSb}== 0    Set Variable            ${ETCdiagResponse.OverCurrent}
    ...                     ELSE IF                 ${MSb}== 1 and ${LSb}== 1    Set Variable            ${ETCdiagResponse.LoadShort}
    ...                     ELSE                    Set Variable            ${ETCdiagResponse.NoOverCurrent}
    [Return]                ${dummy}

CalculateRealAnalogValue
    [Arguments]             ${LSB}                  ${MSB}                  ${R_DOWN}               ${R_UP}
    ${dummy}                Evaluate                ((${LSB} + ${MSB} * ${256}) / ${16383} * ${5.0}) * (${R_DOWN} + ${R_UP}) / ${R_DOWN}
    [Return]                ${dummy}

EXPreadDigital
    [Arguments]             ${in}
    Return From Keyword If    ${in} > ${17}           or                      ${in} < ${0}
    ${Ret}                  Din Get                 ${in}
    Run Keyword If          ${ret}[0] == ${-1}      fail                    Error_Mng               ${CMD_INDICES.EXP_DinGet}    ${ret}[1]
    [Return]                ${ret}[1]

EXPwriteDigital
    [Arguments]             ${in}                   ${state}
    Return From Keyword If    ${in} > ${15}           or                      ${in} < ${0}
    ${Ret}                  Do Set                  ${in}                   ${state}
    # LOG TO CONSOLE          ${Ret}
    Run Keyword If          ${ret}[0] != ${0}       fail                    Error_Mng               ${CMD_INDICES.EXP_DoSet}    ${ret}[1]
    [Return]                ${ret}[1]

EXPFrequencyWrite
    [Arguments]             ${in}                   ${freq}                 ${dutyCycle}
    Return From Keyword If    ${in} > ${5}            or                      ${in} < ${1}
    ${Ret}                  Fo Set                  ${in}                   ${freq}                 ${dutyCycle}
    # LOG TO CONSOLE          ${Ret}
    Run Keyword If          ${ret}[0] != ${0}       fail                    Error_Mng               ${CMD_INDICES.EXP_FoSet}    ${ret}[1]
    [Return]                ${ret}[1]

EXPanalogWrite
    [Arguments]             ${in}                   ${Volt}
    Return From Keyword If    ${in} > ${15}           or                      ${in} < ${0}
    ${Ret}                  Ao Set                  ${in}                   ${Volt}
    Run Keyword If          ${ret}[0] != ${0}       fail                    Error_Mng               ${CMD_INDICES.EXP_AoSet}    ${ret}[1]
    [Return]                ${ret}[1]

Convert to ASCII
    [Arguments]             ${input}
    ${inputs_count} =       Get Length              ${input}
    ${output}               Create List
    FOR                     ${i}                    IN RANGE                ${inputs_count}
        Run Keyword If          ${input[${i}]}== 48     Append To List          ${output}               0
        Run Keyword If          ${input[${i}]}== 49     Append To List          ${output}               1
        Run Keyword If          ${input[${i}]}== 50     Append To List          ${output}               2
        Run Keyword If          ${input[${i}]}== 51     Append To List          ${output}               3
        Run Keyword If          ${input[${i}]}== 52     Append To List          ${output}               4
        Run Keyword If          ${input[${i}]}== 53     Append To List          ${output}               5
        Run Keyword If          ${input[${i}]}== 54     Append To List          ${output}               6
        Run Keyword If          ${input[${i}]}== 55     Append To List          ${output}               7
        Run Keyword If          ${input[${i}]}== 56     Append To List          ${output}               8
        Run Keyword If          ${input[${i}]}== 57     Append To List          ${output}               9
        Run Keyword If          ${input[${i}]}== 65     Append To List          ${output}               A
        Run Keyword If          ${input[${i}]}== 66     Append To List          ${output}               B
        Run Keyword If          ${input[${i}]}== 67     Append To List          ${output}               C
        Run Keyword If          ${input[${i}]}== 68     Append To List          ${output}               D
        Run Keyword If          ${input[${i}]}== 69     Append To List          ${output}               E
        Run Keyword If          ${input[${i}]}== 70     Append To List          ${output}               F
        Run Keyword If          ${input[${i}]}== 71     Append To List          ${output}               G
        Run Keyword If          ${input[${i}]}== 72     Append To List          ${output}               H
        Run Keyword If          ${input[${i}]}== 73     Append To List          ${output}               I
        Run Keyword If          ${input[${i}]}== 74     Append To List          ${output}               J
        Run Keyword If          ${input[${i}]}== 75     Append To List          ${output}               K
        Run Keyword If          ${input[${i}]}== 76     Append To List          ${output}               L
        Run Keyword If          ${input[${i}]}== 77     Append To List          ${output}               M
        Run Keyword If          ${input[${i}]}== 78     Append To List          ${output}               N
        Run Keyword If          ${input[${i}]}== 79     Append To List          ${output}               O
        Run Keyword If          ${input[${i}]}== 80     Append To List          ${output}               P
        Run Keyword If          ${input[${i}]}== 81     Append To List          ${output}               Q
        Run Keyword If          ${input[${i}]}== 82     Append To List          ${output}               R
        Run Keyword If          ${input[${i}]}== 83     Append To List          ${output}               S
        Run Keyword If          ${input[${i}]}== 84     Append To List          ${output}               T
        Run Keyword If          ${input[${i}]}== 85     Append To List          ${output}               U
        Run Keyword If          ${input[${i}]}== 86     Append To List          ${output}               V
        Run Keyword If          ${input[${i}]}== 87     Append To List          ${output}               W
        Run Keyword If          ${input[${i}]}== 88     Append To List          ${output}               X
        Run Keyword If          ${input[${i}]}== 89     Append To List          ${output}               Y
        Run Keyword If          ${input[${i}]}== 76     Append To List          ${output}               Z
        Run Keyword If          ${input[${i}]}== 45     Append To List          ${output}               -
        Run Keyword If          ${input[${i}]}== 32     Append To List          ${output}               ' '
    END
    [Return]                ${output}

TheLastOfTheMohicans
    [Arguments]             ${AllDataToOneExcel}    ${ExcelFileNumber}      ${DateTime}             ${retEXP_HSCAN}
    ${ExcelFileNumber}      Evaluate                ${ExcelFileNumber} + ${1}
    ${ExcelFileName}        Catenate                ${Test_Name}            ${DateTime}             part                    ${ExcelFileNumber} (Interuptted Save!)
    ${ExcelParams}          Create Excel File       ${sResultFilePath}      ${ExcelFileName}        ${Test_Name}
    ${Ret_Val}              Run Keyword If          ${retEXP_HSCAN} == ${null}    Write Matrix To Row     ${sResultFilePath}
    ...                     ${ExcelFileName}        ${Test_Name}       ${AllDataToOneExcel}    row=${1}    col=${1}
    ${AllDataToOneExcel}    Run Keyword If          ${retEXP_HSCAN} == ${null}    Set Variable            ${null}

SendGiveDatafromECU
    [Arguments]             ${DataToSend}
    ${WhichData}            Copy List               ${DataToSend}
    ${i}                    Set Variable            ${0}
    ${WhichData}            CalculateCheckSum       ${WhichData}
    FOR                     ${i}                    IN RANGE                ${MaxResponeseWaitingIteration}
        ${retEXP_HSCAN}         EXP_HSCAN.Diag          ${ID.ECMT}              ${ID.DIAG}              ${WhichData}
        Run Keyword If          ${retEXP_HSCAN} == ${null}    Fail                    \nNull error
        Exit For Loop If        ${retEXP_HSCAN}[0] != ${1}
    END
    Set Global Variable     ${retEXP_HSCAN}         ${retEXP_HSCAN}
    Run Keyword If          ${retEXP_HSCAN}[0] == ${-1}    Log To Console          Reading CAN Failed: Iteration = ${i}
    # ${Retdata}            Run Keyword If          ${retEXP_HSCAN}[0] == ${-1}    Set Variable            ${False}
    ${Retdata}              Run Keyword If          ${retEXP_HSCAN}[0] == ${-1}    Set Variable            ${0}
    Run Keyword If          ${retEXP_HSCAN}[0] == ${1}    Log To Console          No response from ECU: Iteration = ${i}
    Run Keyword If          ${retEXP_HSCAN}[0] == ${0} and ${retEXP_HSCAN}[1][3] == 179    Log To Console          LENGTH_ERROR
    # ${Retdata}            Run Keyword If          ${retEXP_HSCAN}[0] == ${0} and ${retEXP_HSCAN}[1][3] == 160    Set Variable    ${retEXP_HSCAN}    ELSE    ${False}
    ${Retdata}              Run Keyword If          ${retEXP_HSCAN}[0] == ${0} and ${retEXP_HSCAN}[1][3] == 160    Set Variable    ${retEXP_HSCAN}    ELSE    Set Variable    ${0}
    Run Keyword If          ${retEXP_HSCAN}[0] == ${-1}    Set Global Variable     ${CAN_FLAG}             ${1}
    [Return]                ${Retdata}
