import pandas as pd
import xlrd
import streamlit as st

df = pd.read_excel(r'Report.xls', sheet_name='Data')

st.set_page_config(layout='wide')
st.header('Yuming reports')
st.subheader('Create report for concrete serial number:')
serial_number = st.text_input('Serial number:')
sn_search = st.button('Search SN')


if sn_search:
    filter_df = df.loc[(df['AP60ST3scaner code'].str.contains(serial_number))]
    st.write(filter_df)
    if not filter_df.empty:
        total_status = filter_df.iloc[0]['Total Result']

        col60, col90, col95, col140, col150, col155, col160 = st.columns(7)


        with col60:
            ap60tr = filter_df.iloc[0]['AP60Total Result']
            ap60pp = filter_df.iloc[0]['AP60 Part Position']
            ap60ap = filter_df.iloc[0]['AP60 Air Pressure']
            ap60st1r = filter_df.iloc[0]['AP60 ST1 Result']
            ap60camera = filter_df.iloc[0]['AP60 Camera Result']
            ap60iv1 = filter_df.iloc[0]['AP60 IV Tool 1 Result']
            ap60iv2 = filter_df.iloc[0]['AP60 IV Tool 2 Result']
            ap60iv3 = filter_df.iloc[0]['AP60 IV Tool 3 Result']
            ap60laser = filter_df.iloc[0]['AP60 Laser Speed']
            ap60power = filter_df.iloc[0]['AP60 Laser Power']
            ap60st3 = filter_df.iloc[0]['AP60 ST3 Result']
            ap60date = filter_df.iloc[0]['AP60 Date/Time']
            if ap60tr == 'NOK':
                st.error('AP 60 is NOK')
            else:
                st.success('AP60 is OK')

            st.write(f'Total Status: {ap60tr}')
            st.write(f'Part Position: {ap60pp}')
            st.write(f'Air Pressure: {ap60ap}')
            st.write(f'ST1 Result: {ap60st1r}')
            st.write(f'Camera check: {ap60camera}')
            st.write(f'IV Tool 1 Result: {ap60iv1}')
            st.write(f'IV Tool 2 Result: {ap60iv2}')
            st.write(f'IV Tool 3 Result: {ap60iv3}')
            st.write(f'Laser Speed: {ap60laser}')
            st.write(f'Laser Power: {ap60power}')
            st.write(f'ST3 Result: {ap60st3}')
            st.write(f'Date and time: {ap60date}')

        with col90:
            ap90tr = filter_df.iloc[0]['AP90 Total Result']
            ap90st1 = filter_df.iloc[0]['AP90 ST1 Result']
            ap90h1 = filter_df.iloc[0]['AP90 Height Result 1']
            ap90maxh1 = filter_df.iloc[0]['AP90 MAX Height 1(mm)']
            ap90hval1 = filter_df.iloc[0]['AP90 Height Value 1(mm)']
            ap90minh1 = filter_df.iloc[0]['AP90 MIN Height 1(mm)']
            ap90h2 = filter_df.iloc[0]['AP90 Height Result 2']
            ap90maxh2 = filter_df.iloc[0]['AP90 MAX Height 2(mm)']
            ap90hval2 = filter_df.iloc[0]['AP90 Height Value 2(mm)']
            ap90minh2 = filter_df.iloc[0]['AP90 MIN Height 2(mm)']
            ap90weldtin1 = filter_df.iloc[0]['AP90 Weld 1 Tin Content 1 (mm)']
            ap90weldtin12 = filter_df.iloc[0]['AP90 Weld 1 Tin Content 1 (mm)2']
            ap90weldtemp1 = filter_df.iloc[0]['AP90 Weld 1 Temp (℃)']
            ap90weldtin2 = filter_df.iloc[0]['AP90 Weld 2 Tin Content 1 (mm)']
            ap90weldtin22 = filter_df.iloc[0]['AP90 Weld 2 Tin Content 2 (mm)']
            ap90weldtemp2 = filter_df.iloc[0]['AP90 Weld 2 Temp (℃)']
            ap90date = filter_df.iloc[0]['AP90 Date/Time']

            if ap90tr == 'NOK':
                st.error('AP 90 is NOK')
            else:
                st.success('AP90 is OK')

            st.write(f'Total Status: {ap90tr}')
            st.write(f'ST1 Result: {ap90st1}')
            st.write(f'Height Result 1: {ap90h1} mm')
            st.write(f'MAX heigh 1: {ap90maxh1} mm')
            st.write(f'Heigh Value 1: {ap90hval1} mm')
            st.write(f'MIN Heigh 1: {ap90minh1} mm')
            st.write(f'Heigh Result 2: {ap90h2} mm')
            st.write(f'MAX heigh 2: {ap90maxh2} mm')
            st.write(f'Heigh Value 2: {ap90hval2}')
            st.write(f'MIN heigh 2: {ap90minh2}')
            st.write(f'Weld 1 tin content: {ap90weldtin1} mm')
            st.write(f'Weld 1 tin content: {ap90weldtin12} mm2')
            st.write(f'Weld 1 temperature: {ap90weldtemp1} deg C')
            st.write(f'Weld 2 tin content: {ap90weldtin2} mm')
            st.write(f'Weld 2 tin content: {ap90weldtin22} mm2')
            st.write(f'Weld 2 temperature: {ap90weldtemp2} deg C')
            st.write(f'Date and time: {ap90date}')

        with col95:
            ap95tr = filter_df.iloc[0]['AP95 Total Result']
            ap95st1 = filter_df.iloc[0]['AP95 ST1 Result']
            ap95st2 = filter_df.iloc[0]['AP95 ST2 Result']
            ap95cno = filter_df.iloc[0]['AP95 Camera No.']
            ap95cres = filter_df.iloc[0]['AP95 Camera Result']
            ap95w1res = filter_df.iloc[0]['AP95 Weld 1 Result']
            ap95w2res = filter_df.iloc[0]['AP95 Weld 2 Result']
            ap95lrres = filter_df.iloc[0]['AP95 Left/Right Result']
            ap95flip = filter_df.iloc[0]['AP95 180°Flip Result']
            ap95gap = filter_df.iloc[0]['AP95 Gap Result']
            ap95st3r = filter_df.iloc[0]['AP95 ST3 Result']
            ap95st3mks1res1 = filter_df.iloc[0]['AP95 ST3 MKS 1 Result 1']
            ap95st3mks1target1 = filter_df.iloc[0]['AP95 ST3 MKS1 Target Value 1 (Ω)']
            ap95st3mks1test1 = filter_df.iloc[0]['AP95 ST3 MKS1 Test Value 1 (Ω)']
            ap95st3mks1diff1 = filter_df.iloc[0]['AP95 ST3 MKS 1 Difference 1 (%)']
            ap95st3mks1res2 = filter_df.iloc[0]['AP95 ST3 MKS 1 Result 2']
            ap95st3mks1target2 = filter_df.iloc[0]['AP95 ST3 MKS 1 Target Value 2 (Ω)']
            ap95st3mks1test2 = filter_df.iloc[0]['AP95 ST3 MKS 1 Test Value 2 (Ω)']
            ap95st3mks1diff2 = filter_df.iloc[0]['AP95 ST3 MKS 1 Difference 2 (%)']
            ap95st3mks2res1 = filter_df.iloc[0]['AP95 ST3 MKS 2 Result 1']
            ap95st3mks2target1 = filter_df.iloc[0]['AP95 ST3 MKS 2 Target Value 1 (Ω)']
            ap95st3mks2test1 = filter_df.iloc[0]['AP95 ST3 MKS 2 Test Value 1 (Ω)']
            ap95st3mks2diff1 = filter_df.iloc[0]['AP95 ST3 MKS 2 Difference 1 (%)']
            ap95st3mks2res2 = filter_df.iloc[0]['AP95 ST3 MKS 2 Result 2']
            ap95st3mks2target2 = filter_df.iloc[0]['AP95 ST3 MKS 2 Target Value 2 (Ω)']
            ap95st3mks2test2 = filter_df.iloc[0]['AP95 ST3 MKS 2 Test Value 2 (Ω)']
            ap95st3mks2diff2 = filter_df.iloc[0]['AP95 ST3 MKS 2 Difference 2 (%)']
            ap95mr = filter_df.iloc[0]['AP95 ST3 Mark Result']
            ap95ap = filter_df.iloc[0]['AP95 Air Pressure']
            ap95date = filter_df.iloc[0]['AP95 Date/Time']

            if ap95tr == 'NOK':
                st.error('AP95 is NOK')
            else:
                st.success('AP95 is OK')

            st.write(f'Total Result: {ap95tr}')
            st.write(f'ST1 Result: {ap95st1}')
            st.write(f'ST2 Result: {ap95st2}')
            st.write(f'Camera No.: {ap95cno}')
            st.write(f'Camera Result: {ap95cres}')
            st.write(f'Weld 1 Result: {ap95w1res}')
            st.write(f'Weld 2 Result: {ap95w2res}')
            st.write(f'Left/Right Result: {ap95lrres}')
            st.write(f'180° Flip Result: {ap95flip}')
            st.write(f'Gap Result: {ap95gap}')
            st.write(f'ST3 Result: {ap95st3r}')
            st.write(f'ST3 MKS 1 Result 1: {ap95st3mks1res1}')
            st.write(f'ST3 MKS1 Target Value 1 (Ω): {ap95st3mks1target1}')
            st.write(f'ST3 MKS1 Test Value 1 (Ω): {ap95st3mks1test1}')
            st.write(f'ST3 MKS 1 Difference 1 (%): {ap95st3mks1diff1}')
            st.write(f'ST3 MKS 1 Result 2: {ap95st3mks1res2}')
            st.write(f'ST3 MKS 1 Target Value 2 (Ω): {ap95st3mks1target2}')
            st.write(f'ST3 MKS 1 Test Value 2 (Ω): {ap95st3mks1test2}')
            st.write(f'ST3 MKS 1 Difference 2 (%): {ap95st3mks1diff2}')
            st.write(f'ST3 MKS 2 Result 1: {ap95st3mks2res1}')
            st.write(f'ST3 MKS 2 Target Value 1 (Ω): {ap95st3mks2target1}')
            st.write(f'ST3 MKS 2 Test Value 1 (Ω): {ap95st3mks2test1}')
            st.write(f'ST3 MKS 2 Difference 1 (%): {ap95st3mks2diff1}')
            st.write(f'ST3 MKS 2 Result 2: {ap95st3mks2res2}')
            st.write(f'ST3 MKS 2 Target Value 2 (Ω): {ap95st3mks2target2}')
            st.write(f'ST3 MKS 2 Test Value 2 (Ω): {ap95st3mks2test2}')
            st.write(f'ST3 MKS 2 Difference 2 (%): {ap95st3mks2diff2}')
            st.write(f'ST3 Mark Result: {ap95mr}')
            st.write(f'Air Pressure: {ap95ap}')
            st.write(f'Date/Time: {ap95date}')

        with col140:
            ap140tr = filter_df.iloc[0]['AP140 Result']
            ap140ap = filter_df.iloc[0]['AP140 Air Pressure']
            ap140st1r = filter_df.iloc[0]['AP140 ST 1 Result']
            ap140hr = filter_df.iloc[0]['AP140 Height Result']
            ap140hv = filter_df.iloc[0]['AP140 Height Value(mm)']
            ap140maxh = filter_df.iloc[0]['AP140 MAX Height (mm)']
            ap140minh = filter_df.iloc[0]['AP140 MIN Height (mm)']
            ap140ivres = filter_df.iloc[0]['AP140 IV Resistor Result']
            ap140weld1temp = filter_df.iloc[0]['AP140 Weld 1 Temp (℃)']
            ap140weld1tin1 = filter_df.iloc[0]['AP140 Weld 1 Tin Content 1 (mm)']
            ap140weld1tin2 = filter_df.iloc[0]['AP140 Weld 1 Tin Content 2 (mm)']
            ap140weld2temp = filter_df.iloc[0]['AP140 Weld 2 Temp (℃)']
            ap140weld2tin1 = filter_df.iloc[0]['AP140 Weld 2 Tin Content 1 (mm)']
            ap140weld2tin2 = filter_df.iloc[0]['AP140 Weld 2 Tin Content 2 (mm)']
            ap140date = filter_df.iloc[0]['AP140 Date/Time']

            if ap140tr == 'NOK':
                st.error('AP140 is NOK')
            else:
                st.success('AP140 is OK')

            st.write(f'Result: {ap140tr}')
            st.write(f'Air Pressure: {ap140ap}')
            st.write(f'ST 1 Result: {ap140st1r}')
            st.write(f'Height Result: {ap140hr}')
            st.write(f'Height Value(mm): {ap140hv}')
            st.write(f'MAX Height (mm): {ap140maxh}')
            st.write(f'MIN Height (mm): {ap140minh}')
            st.write(f'IV Resistor Result: {ap140ivres}')
            st.write(f'Weld 1 Temp (℃): {ap140weld1temp}')
            st.write(f'Weld 1 Tin Content 1 (mm): {ap140weld1tin1}')
            st.write(f'Weld 1 Tin Content 2 (mm): {ap140weld1tin2}')
            st.write(f'Weld 2 Temp (℃): {ap140weld2temp}')
            st.write(f'Weld 2 Tin Content 1 (mm): {ap140weld2tin1}')
            st.write(f'Weld 2 Tin Content 2 (mm): {ap140weld2tin2}')
            st.write(f'Date/Time: {ap140date}')

        with col150:
            ap150tr = filter_df.iloc[0]['AP150 Total Result']
            ap150ap = filter_df.iloc[0]['AP150 Air Pressure']
            ap150weld1tin1 = filter_df.iloc[0]['AP150 Weld 1 Tin Content 1 (mm)']
            ap150weld1tin2 = filter_df.iloc[0]['AP150 Weld 1 Tin Content 2 (mm)']
            ap150weld1temp = filter_df.iloc[0]['AP150 Weld 1 Temp (℃)']
            ap150weld2tin1 = filter_df.iloc[0]['AP150 Weld 2 Tin Content 1 (mm)']
            ap150weld2tin2 = filter_df.iloc[0]['AP150 Weld 2 Tin Content 2 (mm)']
            ap150weld2temp = filter_df.iloc[0]['AP150 Weld 2 Temp (℃)']
            ap150date = filter_df.iloc[0]['AP150 Date/Time']

            if ap150tr == 'NOK':
                st.error('AP150 is NOK')
            else:
                st.success('AP150 is OK')

            st.write(f'Total Result: {ap150tr}')
            st.write(f'Air Pressure: {ap150ap}')
            st.write(f'Weld 1 Tin Content 1 (mm): {ap150weld1tin1}')
            st.write(f'Weld 1 Tin Content 2 (mm): {ap150weld1tin2}')
            st.write(f'Weld 1 Temp (℃): {ap150weld1temp}')
            st.write(f'Weld 2 Tin Content 1 (mm): {ap150weld2tin1}')
            st.write(f'Weld 2 Tin Content 2 (mm): {ap150weld2tin2}')
            st.write(f'Weld 2 Temp (℃): {ap150weld2temp}')
            st.write(f'Date/Time: {ap150date}')

        with col155:
            sa150soldtime1 = filter_df.iloc[0]['SA150SendSolderingTime']
            sa150soldtime2 = filter_df.iloc[0]['SA150SolderingTime']

            st.write(f'Send soldering time: {sa150soldtime1}')
            st.write(f'Soldering time: {sa150soldtime2}')

        with col160:
            ap160tr = filter_df.iloc[0]['AP160 Total Result']
            ap160camera = filter_df.iloc[0]['AP160 Camera Total Result']
            ap160r1w1 = filter_df.iloc[0]['AP160 Resistor 1 Weld 1 result']
            ap160r1w2 = filter_df.iloc[0]['AP160 Resistor 1 Weld 2 result']
            ap160r2w1 = filter_df.iloc[0]['AP160 Resistor 2 Weld 1 result']
            ap160r2w2 = filter_df.iloc[0]['AP160 Resistor 2 Weld 2 result']
            ap160pyall = filter_df.iloc[0]['AP160 Poka-Yoke All Result']
            pa160pybottom = filter_df.iloc[0]['AP160 Poka-Yoke Bottom Result']
            ap160st3r1 = filter_df.iloc[0]['AP160 ST3 Resistor 1 Result']
            ap160st3target1 = filter_df.iloc[0]['AP160 ST3 Resistor 1 Target Value 1 (Ω)']
            ap160st3stest1 = filter_df.iloc[0]['AP160 ST3 Resistor 1 Test Value 1 (Ω)']
            ap160st3diff1 = filter_df.iloc[0]['AP160 ST3 Resistor 1 Difference 1 (%)']
            ap160st3r2 = filter_df.iloc[0]['AP160 ST3 Resistor 2 Result']
            ap160st3target2 = filter_df.iloc[0]['AP160 ST3 Resistor 2 Target Value 1 (Ω)']
            ap160st3stest2 = filter_df.iloc[0]['AP160 ST3 Resistor 2 Test Value 1 (Ω)']
            ap160st3diff2 = filter_df.iloc[0]['AP160 ST3 Resistor 2 Difference 1 (%)']
            ap160mr = filter_df.iloc[0]['AP160 Mark Result 1']
            ap160ap = filter_df.iloc[0]['AP160 Air Pressure']
            ap160date = filter_df.iloc[0]['AP160 Date/Time']

            if ap160tr == 'NOK':
                st.error('AP160 is NOK')
            else:
                st.success('AP160 is OK')

            st.write(f'Total Result: {ap160tr}')
            st.write(f'Camera Total Result: {ap160camera}')
            st.write(f'Resistor 1 Weld 1 result: {ap160r1w1}')
            st.write(f'Resistor 1 Weld 2 result: {ap160r1w2}')
            st.write(f'Resistor 2 Weld 1 result: {ap160r2w1}')
            st.write(f'Resistor 2 Weld 2 result: {ap160r2w2}')
            st.write(f'Poka-Yoke All Result: {ap160pyall}')
            st.write(f'Poka-Yoke Bottom Result: {pa160pybottom}')
            st.write(f'ST3 Resistor 1 Result: {ap160st3r1}')
            st.write(f'ST3 Resistor 1 Target Value 1 (Ω): {ap160st3target1}')
            st.write(f'ST3 Resistor 1 Test Value 1 (Ω): {ap160st3stest1}')
            st.write(f'ST3 Resistor 1 Difference 1 (%): {ap160st3diff1}')
            st.write(f'ST3 Resistor 2 Result: {ap160st3r2}')
            st.write(f'ST3 Resistor 2 Target Value 1 (Ω): {ap160st3target2}')
            st.write(f'ST3 Resistor 2 Test Value 1 (Ω): {ap160st3stest2}')
            st.write(f'ST3 Resistor 2 Difference 1 (%): {ap160st3diff2}')
            st.write(f'Mark Result 1: {ap160mr}')
            st.write(f'Air Pressure: {ap160ap}')
            st.write(f'Date/Time: {ap160date}')


    else:
        st.write('### OK')
else:
    st.error('Serial number is not in database.')
