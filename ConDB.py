import cx_Oracle
import pandas as pd


class DB:
    def __init__(self, first_day, last_day):
        self.first_day = first_day
        self.last_day = last_day
        


    @staticmethod
    def conn():
        dsn = cx_Oracle.makedsn("10.40.3.10", 1521, service_name="f3ipro")
        connection = cx_Oracle.connect(user=r"focco_consulta", password=r'consulta3i08', dsn=dsn, encoding="UTF-8")
        cur = connection.cursor()
        return cur, connection

    def lib_pcp(self):
        cur, con = self.conn()
        cur.execute(
            r"SELECT ROUND(SUM(TDE.QTDE),2)AS LIB_PCP, TOR.DT_LIBERACAO "
            r"FROM FOCCO3I.TORDENS TOR "
            r"INNER JOIN FOCCO3I.TDEMANDAS TDE ON TDE.ORDEM_ID = TOR.ID "
            r"INNER JOIN FOCCO3I.TITENS_PLANEJAMENTO TPL ON TPL.ID = TDE.ITPL_ID "
            r"INNER JOIN FOCCO3I.TITENS_ESTOQUE EST ON EST.ITEMPR_ID = TPL.ITEMPR_ID "
            r"INNER JOIN (SELECT * FROM FOCCO3I.TORDENS_ROT ROT "
            r"INNER JOIN FOCCO3I.TORDENS_CON COM ON COM.TORDEN_ROT_ID = ROT.ID) DASDA ON DASDA.ORDEM_ID = TOR.ID "
            r"INNER JOIN FOCCO3I.TCENTROS_TRAB TRA ON TRA.ID = DASDA.CENTR_TRAB_ID "
            r"WHERE EST.ALMOX_ID IN (590,591) "
            r"AND TOR.DT_LIBERACAO BETWEEN TO_DATE ('" + str(self.first_day) + "', 'DD/MM/YY') AND TO_DATE ('" + str(self.last_day) + "', 'DD/MM/YY') "
            r"AND TPL.COD_ITEM <> 55955 "
            r"AND TRA.ID IN (53, 40) "
            r"GROUP BY TOR.DT_LIBERACAO "
            r"ORDER BY 2 "
        )
        peso_lib_pcp = cur.fetchall()
        peso_lib_pcp = pd.DataFrame(peso_lib_pcp, columns=['PESO_LIB_PCP', 'DIA'])
        #print('peso_lib_pcp', peso_lib_pcp)
        return round(peso_lib_pcp['PESO_LIB_PCP'].sum(), 2)

    def mp_consumida(self):
        cur, con = self.conn()
        con.commit()
        cur.execute(
            r"SELECT  ROUND(SUM (MOV.QTDE),2) AS MP_CONSUM,  MOV.DT "
            r"FROM FOCCO3I.TMOV_ESTQ MOV "
            r",FOCCO3I.TITENS_ESTOQUE EST "
            r",FOCCO3I.TGRP_CLAS_ITE CLA "
            r",FOCCO3I.TITENS TIT "
            r",FOCCO3I.TTP_MOV_ESTQ TTP "
            r"WHERE MOV.ITESTQ_ID = EST.ID "
            r"AND EST.GRP_CLAS_ID = CLA.ID "
            r"AND EST.COD_ITEM = TIT.COD_ITEM "
            r"AND MOV.TMVES_ID = TTP.ID "
            r"AND TTP.SIGLA IN ('REP', 'REQ', 'AJ1') "
            r"AND MOV.ALMOX_ID_ORIG in (590,591) "
            r"AND MOV.DT BETWEEN TO_DATE ('" + str(self.first_day) + "', 'DD/MM/YY') AND TO_DATE ('" + str(self.last_day) + "', 'DD/MM/YY') "
            r"GROUP BY MOV.DT "
            r"ORDER BY 2 "
        )
        mp_comsumida = cur.fetchall()
        mp_comsumida = pd.DataFrame(mp_comsumida, columns=['PESO_MP_CONSUMIDA', 'DIA'])
        return round(mp_comsumida['PESO_MP_CONSUMIDA'].sum(), 2)

    def peso_car_fechado(self):
        cur, con = self.conn()
        con.commit()
        cur.execute(
            r"SELECT ROUND (SUM(ITPDV.QTDE*ITPDV.PESO_LIQ),2) AS PESO_CAR_FECHADO, CAR.DATA_FIM "
            r"FROM FOCCO3I.TSRENGENHARIA_CARREGAMENTOS CAR "
            r"INNER JOIN FOCCO3I.TSRENGENHARIA_CARREGAMENTOS_IT CARIT ON CARIT.SR_CARREG_ID = CAR.ID "
            r"INNER JOIN FOCCO3I.TITENS_PDV ITPDV                     ON ITPDV.ID = CARIT.ITPDV_ID "
            r"INNER JOIN FOCCO3I.TITENS_COMERCIAL CO                  ON CO.ID = ITPDV.ITCM_ID "
            r"INNER JOIN FOCCO3I.TITENS_ENGENHARIA ENG                ON ENG.ITEMPR_ID = CO.ITEMPR_ID "
            r"WHERE ENG.TP_ITEM = 'F' "
            r"AND CAR.SITUACAO = 'F' "
            r"AND CAR.CARREGAMENTO > 264400 "
            r"AND CAR.DATA_FIM BETWEEN TO_DATE ('" + str(self.first_day) + "', 'DD/MM/YY') AND TO_DATE ('" + str(self.last_day) + "', 'DD/MM/YY') "
            r"GROUP BY CAR.DATA_FIM "
            r"ORDER BY 2 "
        )
        peso_car_fechado = cur.fetchall()
        peso_car_fechado = pd.DataFrame(peso_car_fechado, columns=['PESO_CAR_FECHADO', 'DIA'])
        return round(peso_car_fechado['PESO_CAR_FECHADO'].sum(), 2)

    def horas_elev(self):
        cur, con = self.conn()
        con.commit()
        cur.execute(
            r"SELECT  round((sum(TMOV.tempo)/60),2)AS horas_elev, EXTRACT(MONTH FROM TMOV.DT_APONT) AS MES "
            r"FROM FOCCO3I.TORDENS_MOVTO TMOV "
            r",FOCCO3I.TORDENS_ROT ROT "
            r",FOCCO3I.TORDENS TOR "
            r",FOCCO3I.TSRENG_ORDENS_PROJ PRJ "
            r"WHERE TMOV.TORDEN_ROT_ID = ROT.ID "
            r"AND ROT.ORDEM_ID = TOR.ID "
            r"AND TOR.ID = PRJ.ORDEM_ID "
            r"AND PRJ.FUNC_ID = 1461 "
            r"AND TMOV.DT_APONT BETWEEN TO_DATE('01/01/' || EXTRACT(YEAR FROM SYSDATE), 'DD/MM/YY') AND SYSDATE "
            r"GROUP BY EXTRACT(MONTH FROM TMOV.DT_APONT) "
        )
        horas_elev = cur.fetchall()
        horas_elev = pd.DataFrame(horas_elev, columns=['horas_elev', 'MES'])
        this_month = int(self.first_day[6:7])
        horas_elev = horas_elev[horas_elev.MES == this_month]
        return horas_elev['horas_elev'].sum()

    def horas_tot(self):
        cur, con = self.conn()
        con.commit()
        cur.execute(
            r"SELECT  round((sum(TMOV.tempo)/60),2) AS TEMP_TOTAL, EXTRACT(month FROM TMOV.DT_APONT) AS MES "
            r"FROM FOCCO3I.TORDENS_MOVTO TMOV "
            r"WHERE TMOV.DT_APONT BETWEEN TO_DATE('01/01/' || EXTRACT(YEAR FROM SYSDATE), 'DD/MM/YY') AND SYSDATE "
            r"GROUP BY EXTRACT(month FROM TMOV.DT_APONT) "
        )
        horas_tot = cur.fetchall()
        horas_tot = pd.DataFrame(horas_tot, columns=['horas_tot', 'MES'])
        this_month = int(self.first_day[6:7])
        horas_tot = horas_tot[horas_tot.MES == this_month]
        return horas_tot['horas_tot'].sum()
