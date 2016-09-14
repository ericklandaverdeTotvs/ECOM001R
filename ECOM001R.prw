/*
+-----------------------------------------------------------------------+
| TOTVS MEXICO SA DE CV - Todos los derechos reservados.                |
|-----------------------------------------------------------------------|
|    Cliente:                                                           |
|    Archivo: ECOM001R.PRW                                              |
|   Objetivo: Impresión de Pedido de Compra Modelo 1.                   |
| Responable: Filiberto Pérez                                           |
|      Fecha: Junio del 2014                                            |
+-----------------------------------------------------------------------+
*/
#include "stdwin.ch"
#include "Fileio.ch" 
#INCLUDE "RWMAKE.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "PROTHEUS.CH"
#INCLUDE "AP5MAIL.CH"
#INCLUDE "RPTDEF.CH"  
#INCLUDE "FWPrintSetup.ch"
//-----------------------------------------------------------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------------------------------------------------------

User Function ECOM001R()
Local cNomPerg	:="ECOM001R"  
Local cQry0	  	:= ""         
Local dTotal  	:=0
Local cTotal  	:=""                              
Local iMoeda  	:=0     

 GPEPerg(cNomPerg)	
 Pergunte(cNomPerg,.F.)

	@ 200,1 TO 400,377 DIALOG oPedCom TITLE OemToAnsi("Impresión Pedido de Compra")
	@ .5,.5 TO 6,23
	@ 01,001 Say " Este imprime el formato de Pedido de Compra de acuerdo a los     "  
	@ 02,001 Say " parametros informados por el usuario.                            " 
	@ 03,001 Say "                                                                  "
	@ 04,001 Say "                                                                  "
	@ 86,095 BMPBUTTON TYPE 5 ACTION Pergunte(cNomPerg) // Boton de Parametros
	@ 86,125 BMPBUTTON TYPE 01 ACTION RptNota() 	 	// Boton de Generación e Impresion
	@ 86,155 BMPBUTTON TYPE 02 ACTION Close(oPedCom)
	
	Activate Dialog oPedCom Centered
Return
//-----------------------------------------------------------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------------------------------------------------------

Static Function RptNota()
Local ni       		:=0
Local pag      		:=0
Local _nReg    		:=0 
local cCom				:= ""
local i				:=0

Private oPrint
Private oFont
Private nRenIni  		:= 50 
Private _nItem   		:=0
Private nColIni  		:= 20 
Private nLin     		:=nRenIni+800
Private _Subtotal		:=0
Private _nDesc   		:=0    
Private _nIva   		:=0
Private _cusuario 	:=""
Private _caprobador	:=""  
Private _deposito		:="" 
Private _cdocaprob	:=""
private nPagNum 		:= 0   
Private lEmail		:= .f.
Private cFileName		:= ""
     

/*Prepara Caracteristicas de la Letra*/
nHeight   :=15
lBold     := .F.
lUnderLine:= .F.
lPixel    := .T.
lPrint    := .F.

oFont := TFont():New( "Arial",,nHeight,,lBold,,,,,lUnderLine )
oFont1:= TFont():New( "Arial",,10,,.t.,,,,,.f. )
oFont2:= TFont():New( "Arial",,8,,.f.,,,,,.f. )
oFont3:= TFont():New( "Arial",,8,,.f.,,,,,.f. ) 
oFont5:= TFont():New( "Arial",,6,,.f.,,,,,.f. ) 
oFont4:= TFont():New( "Arial",,14,,.f.,,,,,.f. )   //Courier New 

Private oArial08		:= TFont():New("Arial",08,08,,.F.,,,,.T.,.F.)
Private oArial08N		:= TFont():New("Arial",08,08,,.T.,,,,.T.,.F.)
Private oArial09		:= TFont():New("Arial",09,09,,.F.,,,,.T.,.F.)
Private oArial09N		:= TFont():New("Arial",09,09,,.T.,,,,.T.,.F.)
Private oArial10		:= TFont():New("Arial",10,10,,.F.,,,,.T.,.F.)
Private oArial10N		:= TFont():New("Arial",12,12,,.T.,,,,.T.,.F.)
Private oArial11		:= TFont():New("Arial",11,11,,.F.,,,,.T.,.F.)
Private oArial11N		:= TFont():New("Arial",11,11,,.T.,,,,.T.,.F.)
Private oArial14N		:= TFont():New("Arial",14,14,,.T.,,,,.T.,.F.)
Private oLucCon10		:= TFont():New("Lucida Console",10,10,,.F.,,,,.T.,.F.)
Private oLucCon10N	:= TFont():New("Lucida Console",10,10,,.T.,,,,.T.,.F.)

Private	cNumPed  	:= mv_par01			// Numero de Pedido de Compras

dbSelectArea("SC7")//Encabezado de Pedidos
dbsetorder(1)
dbseek(xFilial("SC7")+mv_par01) 

if dbseek(xFilial("SC7")+mv_par01) 
else
   msgalert("El Pedido de compra no existe, Verifique...")
   return
end if   
/*
If MsgYesNo("¿Desea Enviar el Pedido de Compra por Email?") 
	lEmail := .t.
Endif
*/
lViewPDF := !lEmail

_nReg:=0
_usuario:=SC7->C7_USER

cFileName := ALLTRIM(SC7->C7_NUM) + "_Mod1.pdf"
oPrint	:= FWMsPrinter():New(cFileName,6,.T.,,.T.,,,,,,,lViewPDF,)
oPrint:SetResolution()
oPrint:SetPortrait()
oPrint:cPathPDF:= "C:\SPOOL\PedidosCompra\"

While ! eof() .and. SC7->C7_FILIAL==xFilial("SC7") .and. SC7->C7_NUM==mv_par01
      _nReg++
      dbskip()
end

ProcRegua(_nReg) //Inicia barra de procesamiento con numero de registros a procesar
dbseek(xFilial("SC7")+mv_par01)   
      
_deposito := SC7->C7_LOCAL         
_obs:= SC7->C7_OBSE
                
Encabez()

dbSelectArea("SC7")//Encabezado de Pedidos
dbsetorder(1)
dbseek(xFilial("SC7")+mv_par01) 

WHILE !Eof() .AND. xFilial("SC7") == SC7->C7_FILIAL .AND. SC7->C7_NUM == mv_par01

    IncProc("Imprimiendo...")	
	SB1->(DBSETORDER(1))//Catalogo de productos
	SB1->(dbseek(xFILIAL("SB1")+SC7->C7_PRODUTO))
	_cusuario := SC7->C7_USER
	_caprobador := SC7->C7_APROV  
	_cdocaprob := SC7->C7_CONAPRO

	/*Imprime Detalle */
	oPrint:Say( nLin,nColIni , SC7->C7_ITEM           		                 		, oLucCon10, 100)//ITEM
	oPrint:Say( nLin,nColIni+50   , TRANSFORM(SC7->C7_QUANT,"@E 9,999,999")    		, oLucCon10, 100)//Cantidad
	oPrint:Say( nLin,nColIni+300 , SC7->C7_PRODUTO                            		, oLucCon10, 100)//Clave
	oPrint:Say( nLin,nColIni+1400, TRANSFORM(SC7->C7_PRECO,"@E 9,999,999.9999")   	, oLucCon10, 100)//costo unitario
	oPrint:Say( nLin,nColIni+1900, SB1->B1_UM                                		, oLucCon10, 100)//UM
	oPrint:Say( nLin,nColIni+2000, TRANSFORM(SC7->C7_TOTAL, "@E 999,999,999.9999")	, oLucCon10, 100)//importe
	
	ni   :=1
	_nInc:=40
	WHILE ni < len(TRIM(SB1->B1_DESC))
		oPrint:Say( nLin,nColIni+600, SUBSTR(TRIM(SB1->B1_DESC),ni,45), oLucCon10, 100)
		ni  +=45
		nLin+=_nInc
		_nItem++
	ENDDO

	cQry0	:= "SELECT ACB.ACB_DESCRI from "+RetSQLName("ACB")+" ACB, "+RetSQLName("AC9")+" AC9 "
	cQry0	+= "WHERE "
	cQry0	+= "AC9.AC9_CODOBJ=ACB.ACB_CODOBJ "
	cQry0	+= "AND AC9_CODENT='"+SC7->C7_PRODUTO+"' "
	cQry0	+= "AND AC9.D_E_L_E_T_<>'*' "
	cQry0	+= "AND ACB.D_E_L_E_T_<>'*' "	
	
	cVal := GetNextAlias()
	dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQry0),cVal,.T.,.T.)

cObs :=  Alltrim(SC7->C7_OBSE) // Variable que contiene la descriocuion de la pieza					
cCadena:= cObs

nLimite:= 68
nResto:= len(cCadena)//-nLimite
nPosIni:= 1
nPosFin:= 0
nLinDes := 0
nFall	:= 1

IF LEN(cCadena)<68
oPrint:Say(nLin,nColIni+600,cCadena,oLucCon10)
//iNum ++
nLin  += 50			
ENDIF

while nResto >=68
//while len(cCadena) >=55

		nPosFin:= 68
		IF LEN(cCadena)>=68
			nPosFin:= rat(" ",substr(cCadena,nPosIni,nPosFin) )
		ELSE
			nPosFin:= 68 
		ENDIF               
		IF nPosFin == 0
			cImprime:=cCadena
			//msgalert(cImprime) //imprimir
			//oPrint:Say(nLin,,cImprime,oLucCon10)
			oPrint:Say(nLin,nColIni+600,cImprime,oLucCon10)
					//iNum ++
					//nLinDes ++ 
					//nLinDes ++
					//nFall ++
					nLin  += 50 //OJO 
			//BREAK	
		ENDIF
		cImprime:=substr(cCadena,nPosIni,nPosFin)
		
		//oPrint:Say(nLin + (nLinDes * 35),0480,cImprime,oLucCon10)		//Descripcion
		oPrint:Say(nLin + (nLinDes * 35),nColIni+600,cImprime,oLucCon10)		//Descripcion
					//iNum ++
					//nLoop ++
					//nLinDes ++
					//nLinDes ++
					//nFall ++
					nLin  += 50 //OJO
		nResto:= len(cCadena)-nPosIni
		cCadena:=alltrim(substr(cCadena, nPosFin , len(cCadena)))

enddo
//nLinDes ++
//nFall ++


/*	
	if !empty(ALLTRIM(SC7->C7_OBSE))
			nLoop	:= 45  
		   	For i:=1 To Len(SC7->C7_OBSE) Step 110  
	        oPrint:Say(nLin,nColIni+600 ,substr(SC7->C7_OBSE,i,110),oLucCon10, 100)
		  	nLin+=nLoop 
		  	_nItem++
		  	Next i
		    nLin+=_nInc   
	else 
		SB5->(DBSETORDER(1)) 
		SB5->(dbseek(xFILIAL("SB5")+SC7->C7_PRODUTO))
		cCom := Alltrim(SB5->B5_CEME)					

		if !empty(cCom)
			nLoop	:= 45  
		   	For i := 1 To Len(cCom) Step 110  
	        oPrint:Say(nLin,nColIni ,substr(cCom,i,110),oLucCon10, 100)
		  	nLin+=nLoop 
		  	_nItem++
		  	Next i
		    nLin+=_nInc 			  
		else		    
		    oPrint:Say(nLin,nColIni, alltrim(cCom), oLucCon10,100)
	        nLin+=_nInc
	        ( dbSkip() )
	        _nItem++ 
	     endif
	     
	Endif  
*/

	iMoeda:= SC7->C7_MOEDA
	_Subtotal+= SC7->C7_TOTAL
	_nDesc   += SC7->C7_VLDESC
	_nIva+= SC7->C7_VALIMP1
	dTotal := _Subtotal - _nDesc+_nIva
	cTotal := Implet(dTotal,iMoeda)
	
	if _nItem >= 19
		oPrint:Say (nRenIni+2300+200,nColIni,Replicate("_",250),oLucCon10, 100 )
		oPrint:Say (nRenIni+2350+200,nColIni,"OBSERVACIONES: ",oArial10N, 100 )
		oPrint:Say( nRenIni+2330+200,nColIni+300 ,ALLTRIM(MV_PAR02),oArial10N, 100 )
		oPrint:Say (nRenIni+2360+200,nColIni+300 ,ALLTRIM(MV_PAR03),oArial10N, 100 )
		oPrint:Say (nRenIni+2390+200,nColIni+300 ,ALLTRIM(MV_PAR04),oArial10N, 100 )
		oPrint:Say( nRenIni+2420+200,nColIni+300 ,ALLTRIM(MV_PAR05),oArial10N, 100 )
		oPrint:Say (nRenIni+2450+200,nColIni+300 ,ALLTRIM(MV_PAR06),oArial10N, 100 )
		oPrint:Say (nRenIni+2480+200,nColIni+300 ,ALLTRIM(MV_PAR07),oArial10N, 100 )
	
		oPrint:Say (nRenIni+2750+150,nColIni+120 ,Replicate("_",30)              ,oLucCon10, 100 )
		oPrint:Say (nRenIni+2800+150,nColIni+150 ,"ELABORÓ Nombre y Firma"       ,oArial10N, 100 )
		oPrint:Say (nRenIni+2750+150,nColIni+150 , POSICIONE("SY1",3,xfilial("SY1")+_cusuario ,"Y1_NOME"),oArial10N, 100 ) // USUARIO ELABORO
		oPrint:Say (nRenIni+2750+150,nColIni+970 ,Replicate("_",30)              ,oLucCon10, 100 )
		oPrint:Say (nRenIni+2800+150,nColIni+1000,"AUTORIZÓ Nombre y Firma"      ,oArial10N, 100 )
		
		oPrint:Say (nRenIni+2750+150,nColIni+1820,Replicate("_",30)              	,oLucCon10, 100 )
		oPrint:Say (nRenIni+2800+150,nColIni+1850,"Vo. Bo. Dirección"	,oArial10N, 100 )
						
		oPrint:EndPage()
		pag:=pag+1
		
		Encabez()
		_nItem:=0; nLin:=nRenIni+800
	endif                        

	dbselectarea("SC7")
	dbskip()
END

//msgalert(_nItem)

/*IMPRIME Notas del Pedido, SUBTOTALES Y TOTALES */                                
oPrint:Say (nRenIni+2300+200,nColIni     ,Replicate("_",250)             ,oLucCon10, 100 )
oPrint:Say (nRenIni+2350+200,nColIni     ,"OBSERVACIONES: "              ,oArial10N, 100 )
oPrint:Say (nRenIni+2330+200,nColIni+300 ,ALLTRIM(MV_PAR02)              ,oArial10N, 100 )
oPrint:Say (nRenIni+2360+200,nColIni+300 ,ALLTRIM(MV_PAR03)              ,oArial10N, 100 )
oPrint:Say (nRenIni+2390+200,nColIni+300 ,ALLTRIM(MV_PAR04)              ,oArial10N, 100 )
oPrint:Say (nRenIni+2420+200,nColIni+300 ,ALLTRIM(MV_PAR05)              ,oArial10N, 100 )
oPrint:Say (nRenIni+2450+200,nColIni+300 ,ALLTRIM(MV_PAR06)              ,oArial10N, 100 )
oPrint:Say (nRenIni+2480+200,nColIni+300 ,ALLTRIM(MV_PAR07)              ,oArial10N, 100 )

oPrint:Say (nRenIni+2750+150,nColIni+120 ,Replicate("_",30)              ,oLucCon10, 100 )
oPrint:Say (nRenIni+2800+150,nColIni+150 ,"ELABORÓ Nombre y Firma"       ,oArial10N, 100 )
oPrint:Say (nRenIni+2750+150,nColIni+150 , POSICIONE("SY1",3,xfilial("SY1")+_cusuario ,"Y1_NOME"),oArial10N, 100 ) // USUARIO ELABORO
oPrint:Say (nRenIni+2750+150,nColIni+970 ,Replicate("_",30)              ,oLucCon10, 100 )
oPrint:Say (nRenIni+2800+150,nColIni+1000,"AUTORIZÓ Nombre y Firma"      ,oArial10N, 100 )

oPrint:Say (nRenIni+2750+150,nColIni+1820,Replicate("_",30)              	,oLucCon10, 100 )
oPrint:Say (nRenIni+2800+150,nColIni+1850,"Vo. Bo. Dirección"            ,oArial10N, 100 )

oPrint:Say (nRenIni+2350+200,nColIni+1750,"Subtotal: "                     	,oArial10N, 100)
// oPrint:Say( nRenIni+2400+200,nColIni+1750,"Descuento:"                    	,oArial10N, 100)
oPrint:Say( nRenIni+2400+200,nColIni+1750,"IVA:      "                        ,oArial10N, 100) 
oPrint:Say( nRenIni+2450+200,nColIni+1750,"Total:    "                        ,oArial10N, 100)


oPrint:Say (nRenIni+2350+200,nColIni+2130,TRANSFORM(_Subtotal,"999,999,999.9999")         		,oLucCon10N, 100 )
//oPrint:Say (nRenIni+2400+200,nColIni+2130,TRANSFORM(_nDesc   ,"999,999,999.9999")          		,oLucCon10N, 100 )
oPrint:Say (nRenIni+2400+200,nColIni+2130,TRANSFORM(_nIva ,"999,999,999.9999")					,oLucCon10N, 100 )
oPrint:Say (nRenIni+2450+200,nColIni+2130,TRANSFORM(_Subtotal - _nDesc+_nIva,"999,999,999.9999"),oLucCon10N, 100 )

oPrint:Say( 2500,020,ALLTRIM(cTotal),oArial10N, 100)       

oPrint:EndPage()// Salto de pagina
oPrint:preview() //Despliega pantalla de para ver facturas
MS_FLUSH()

oPrint:EndPage()
oPrint:Print()
FreeObj(oPrint)

If lEmail
	PedMail()
Endif

Return
//-----------------------------------------------------------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------------------------------------------------------

static function Encabez()

/*Imprime datos de la empresa*/
oPrint:StartPage() //Inicia Hoja

cFileLogoR:=GetSrvProfString("Startpath","") + "logoFac.jpg"
                                             

oPrint:SayBitmap(nRenIni-20,nColIni+50,cFileLogoR,475,176) // Tem que estar abaixo do RootPath

oPrint:Say( nRenIni,nColIni+700     ,Alltrim(SM0->M0_NOMECOM)												,oArial14N) 
oPrint:Say( nRenIni+70,nColIni+700  ,Alltrim(SM0->M0_ENDCOB) + " " + AllTrim(SM0->M0_COMPCOB)				,oArial10N) 
oPrint:Say( nRenIni+120,nColIni+700 ,AllTrim(SM0->M0_BAIRCOB) + ", " + AllTrim(SM0->M0_CIDCOB)				,oArial10N) 
oPrint:Say( nRenIni+170,nColIni+700 ,AllTrim(SM0->M0_ESTCOB) + " México. C.P. " + AllTrim(SM0->M0_CEPCOB)	,oArial10N) 
oPrint:Say( nRenIni+220,nColIni+700 ,"RFC: " + Alltrim(SM0->M0_CGC) + ". Teléfono: " + Alltrim(SM0->M0_TEL)	,oArial10N) 

/*IMPRIME DATOS DEL PROVEEDOR*/
dbselectarea("SA2")
dbSetOrder(1)
IF SA2->(dbseek(xFILIAL("SA2")+SC7->C7_FORNECE+SC7->C7_LOJA))
	dbselectarea("SYA")
	dbSetOrder(1) 
	
	cNumPro  	:= ALLTRIM(SC7->C7_FORNECE)
	cProNom		:= ALLTRIM(SA2->A2_NOME)
	cProRfc		:= ALLTRIM(SA2->A2_CGC)
	cProCalle	:= ALLTRIM(SA2->A2_END)
	cProNumExt	:= ALLTRIM(SA2->A2_NR_END)
	cProNumInt	:= ALLTRIM(SA2->A2_NROINT)
	cProMun		:= ALLTRIM(SA2->A2_MUN)
	cProCol		:= ALLTRIM(SA2->A2_BAIRRO)
	cProEst		:= AllTrim(POSICIONE("SX5", 1, XFILIAL("SX5") + '12' + SA2->A2_EST, 'SX5->X5_DESCSPA'))
	cProPais	:= AllTrim(POSICIONE("SYA", 1, XFILIAL("SX5") + SA2->A2_PAIS, 'SYA->YA_DESCR'))
	cProCp		:= ALLTRIM(SA2->A2_CEP)    
	cTelP	   	:= ALLTRIM(SA2->A2_TEL)
	cProCont  	:= ALLTRIM(SA2->A2_CONTATO)	
	cProCC    	:= AllTrim(POSICIONE("CTT", 1, XFILIAL("CTT") + SC7->C7_CC, 'CTT->CTT_DESC01'))     
	cProVia		:= ALLTRIM(SC7->C7_XVIA)   
	cProResCom  := ALLTRIM(SC7->C7_RESCOMP)   
	cProNat 	:= ALLTRIM(SC7->C7_NATUREZ)
	
	oPrint:Say( nRenIni+300,50,"PROVEEDOR/TO:  ", oArial10N, 100)/*Nombre */
	oPrint:Say( nRenIni+350,50,"R.F.C:      ", oArial10N, 100)/*R.F.C. */
	oPrint:Say( nRenIni+400,50,"Direccion/Address:  ", oArial10N, 100)/*Direccion, Colonia */
	oPrint:Say( nRenIni+550,50,"Embarcar a/Ship to: ", oArial10N, 100)/*Email, tel y fax */
	oPrint:Say( nRenIni+600,50,"Via/Shipping Instructions: ", oArial10N, 100)
	
	oPrint:Say( nRenIni+300,400,"(" + cNumPro + ") - " + cProNom, oArial10, 100) /*Nombre */ 
	oPrint:Say( nRenIni+350,400,cProRfc	,oArial10, 100)/*R.F.C. */
	oPrint:Say( nRenIni+400,400,cProCalle + " " + cProNumExt + " " + cProNumInt + ", " + cProCol,oArial10, 100) /*Direccion, Colonia */	
	oPrint:Say( nRenIni+450,400,cProMun + ", " + cProEst + " C.P. " + cProCp +  " " + cProPais,oArial10, 100) /*Ciudad, CP */ 
	oPrint:Say( nRenIni+500,400,"TEL. " + cTelP + " FAX. "+SA2->A2_FAX,oArial10, 100) /*Email, tel y fax */ 
	oPrint:Say( nRenIni+550,400,ALLTRIM(MV_PAR08),oArial10) /*Embarcar a-Facturar a*/
	oPrint:Say( nRenIni+600,400,"" + cProVia + " ",oArial10, 160) /*Via-Shipping instructions*/
	//oPrint:Say( nRenIni+650,50,ALLTRIM(MV_PAR09)+ " " + ALLTRIM(MV_PAR10)								,oArial10)
	oPrint:Say( nRenIni+650,50,ALLTRIM(MV_PAR09)	,oArial10)
	oPrint:Say( nRenIni+680,50,ALLTRIM(MV_PAR10)	,oArial10)
	
	nPagNum := nPagNum + 1
	oPrint:Say(0050,2200,"Página: "+Transform(nPagNum,"999"),oArial10N) // Numero de la pagina

	SE4->(dbseek(xfilial("SE4")+SC7->C7_COND))
	
	IF SC7->C7_MOEDA == 2
		//oPrint:Say( nRenIni+650,nColIni+800,"DOCUMENTO EXPRESADO EN DOLARES AMERICANOS", oArial11N, 100)
		//oPrint:Say( nRenIni+600,nColIni+1000,"DOCUMENTO EXPRESADO EN DOLARES AMERICANOS", oArial11N, 100) 
	ENDIF          
	oPrint:Say( nRenIni+250,nColIni+1500,"Responsable compra"					, oArial10N, 100)
	oPrint:Say( nRenIni+300,nColIni+1500,"No. Pedido/Purchase Order:"			, oArial10N, 100) 
	oPrint:Say( nRenIni+350,nColIni+1500,"Fecha / Date:        "				, oArial10N, 100)
	oPrint:Say( nRenIni+400,nColIni+1500,"Cond. Pago /Terms:   "				, oArial10N, 100) 
//	oPrint:Say( nRenIni+450,nColIni+1500,"CENTRO COSTO: "						, oArial10N, 100) 
//	oPrint:Say( nRenIni+500,nColIni+1500,"Contacto/Contact:     "				, oArial10N, 100)
	oPrint:Say( nRenIni+450,nColIni+1500,"Naturaleza: "							, oArial10n,100)
	oPrint:Say( nRenIni+500,nColIni+1500,"Fecha Embarque / Ship Date:        "	, oArial10N, 100) 

	oPrint:Say( nRenIni+250 ,nColIni+2000,cProResCom           		,oArial10)//Responasble de compra           
	oPrint:Say( nRenIni+300 ,nColIni+2000,SC7->C7_NUM             ,oArial10N)//No. Pedido
	oPrint:Say( nRenIni+350 ,nColIni+2000,dtoc(SC7->C7_EMISSAO)	,oArial10N)//Fecha Elaborac.
	oPrint:Say( nRenIni+400 ,nColIni+2000,SE4->E4_DESCRI        	,oArial10)//Fecha Elaborac.
	//oPrint:Say( nRenIni+500 ,nColIni+2000,cProCont           		,oArial10) 
	//oPrint:Say( nRenIni+450 ,nColIni+2000,cProCC           		,oArial10)
	oPrint:Say( nRenIni+450,2010,"" + cProNat + " ",oArial10, 100)  //Naturaleza
	oPrint:Say( nRenIni+500,nColIni+2000,dtoc(SC7->C7_XFEMB)	,oArial10N)//Fecha Embarque-Ship date
	
ENDIF
oPrint:Say (nRenIni+690,nColIni     , Replicate("_",250)            ,oLucCon10, 100)
oPrint:Say( nRenIni+730,nColIni     , "ITEM"             			,oArial10N, 100)
oPrint:Say( nRenIni+730,nColIni+80  , "Cantidad / QTY"             	,oArial10N, 100)
oPrint:Say( nRenIni+730,nColIni+300 , "Clave"                       ,oArial10N, 100)
oPrint:Say( nRenIni+730,nColIni+600 , "Descripción/Description"	    ,oArial10N, 100) 
oPrint:Say( nRenIni+730,nColIni+1400, "Costo Unitario/Unit Price"   ,oArial10N, 100)
oPrint:Say( nRenIni+730,nColIni+1900, "UM"                   	    ,oArial10N, 100)
oPrint:Say( nRenIni+730,nColIni+2000, "Importe/Amount"              ,oArial10N, 100)
oPrint:Say (nRenIni+750,nColIni     , Replicate("_",250)            ,oLucCon10, 100)
return
//-----------------------------------------------------------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------------------------------------------------------

Static Function ImpLet(pTotal,pMoneda)
                            
    If  pMoneda == 1
          _cSimbM := " $ "
          //Sintaxe: Extenso(nValor,lQuantid,nMoeda,cPrefixo,cIdioma,lCent,lFrac)
          _cLin := Extenso(pTotal,.F.,1,,"2",.T.,.T.,.F.,"2")        
           cCentavos := Right(_cLin,9)
          _cLin := "("+Left(_cLin,Len(_cLin)-9)+cCentavos+")"
    
    ElseIF pMoneda == 2
    	  _cSimbM := " USD$ "
          _cLin := Extenso(pTotal,.F.,2,,"2",.T.,.T.,.F.,"2")
          cCentavos := Right(_cLin,8)
          _cLin :="("+ Left(_cLin,Len(_cLin)-8)+cCentavos + ")"
    
    ElseIf pMoneda == 3
           _cSimbM := " € "  
           _cLin := Extenso(pTotal,.F.,3,,"2",.T.,.T.,.F.,"2")
           cCentavos := Right(_cLin,8)
           _cLin :="("+ Left(_cLin,Len(_cLin)-8)+cCentavos + ")"
    Else
          _cSimbM := " USD$ "
          _cLin := Extenso(pTotal,.F.,2,,"2",.T.,.T.,.F.,"2")
          cCentavos := Right(_cLin,8)
          _cLin :="("+ Left(_cLin,Len(_cLin)-8)+cCentavos + ")"
    EndIF
                                      
Return(_cLin)
//-----------------------------------------------------------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------------------------------------------------------     
       
/*BEGINDOC
  Generación del Juego de preguntas---------------------------------------------------------------------
  ENDDOC*/

//-----------------------------------------------------------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------------------------------------------------------
Static Function GPEPerg()
		Local _sAlias := Alias()                              
		Local i := 0
		Local j:= 0
		dbSelectArea("SX1")
		dbSetOrder(1)
		
		cPerg := PADR("ECOM001R",10)
		aRegs:={}//G=Edit S=Texto C=Combo el siguiente parametro es para el Valid

		aAdd(aRegs,{cPerg,"01","¿Cod. Pedido?  ","¿Cod. Pedido?  ","¿Cod. Pedido?  ","MV_CH1","C",06,0,0,"G","","MV_PAR01","","","","","","","","","","","","","","","","","","","","","","","","","SC7"})
		aAdd(aRegs,{cPerg,"02","Observaciones 1","Observaciones 1","Observaciones 1","MV_CH2","C",70,0,0,"G","","MV_PAR02","","","","","","","","","","","","","","","","","","","","","","","","",""})
		aAdd(aRegs,{cPerg,"03","Observaciones 2","Observaciones 2","Observaciones 2","MV_CH3","C",70,0,0,"G","","MV_PAR03","","","","","","","","","","","","","","","","","","","","","","","","",""})
		aAdd(aRegs,{cPerg,"04","Observaciones 3","Observaciones 3","Observaciones 3","MV_CH4","C",70,0,0,"G","","MV_PAR04","","","","","","","","","","","","","","","","","","","","","","","","",""})
		aAdd(aRegs,{cPerg,"05","Observaciones 4","Observaciones 4","Observaciones 4","MV_CH5","C",70,0,0,"G","","MV_PAR05","","","","","","","","","","","","","","","","","","","","","","","","",""})
		aAdd(aRegs,{cPerg,"06","Observaciones 5","Observaciones 5","Observaciones 5","MV_CH6","C",70,0,0,"G","","MV_PAR06","","","","","","","","","","","","","","","","","","","","","","","","",""})
		aAdd(aRegs,{cPerg,"07","Observaciones 6","Observaciones 6","Observaciones 6","MV_CH7","C",70,0,0,"G","","MV_PAR07","","","","","","","","","","","","","","","","","","","","","","","","",""})
		aAdd(aRegs,{cPerg,"08","Facturar a 1   ","Embarcar a     ","Embarcar a    ","MV_CH8","C",99,0,0,"G","","MV_PAR08","","","","","","","","","","","","","","","","","","","","","","","","",""})
		aAdd(aRegs,{cPerg,"09","Facturar a 2   ","Facturar a 2   ","Facturar a 2   ","MV_CH9","C",99,0,0,"G","","MV_PAR09","","","","","","","","","","","","","","","","","","","","","","","","",""})
		aAdd(aRegs,{cPerg,"10","Facturar a 3   ","Facturar a 3   ","Facturar a 3   ","MV_CH10","C",99,0,0,"G","","MV_PAR10","","","","","","","","","","","","","","","","","","","","","","","","",""})
		For i:=1 to Len(aRegs)
			If !dbSeek(cPerg+aRegs[i,2])
				RecLock("SX1",.T.)
				For j:=1 to FCount()
					If j <= Len(aRegs[i])
						FieldPut(j,aRegs[i,j])
					Endif
				Next
				MsUnlock()
			Endif
		Next
		dbSelectArea(_sAlias)
return
//-----------------------------------------------------------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------------------------------------------------------
Static Function PedMail()

Private oMailPC
Private nTarget:=0
Private cFOpen :=""
Private nOpc   := 0
Private bOk    := {||nOpc:=1,oMailPC:End()}
Private bCancel:= {||nOpc:=0,oMailPC:End()} 
Private lCheck1:=.F.
Private lCheck2:=.T.
Private lCheck3:=.f.


_cPara   :=SA2->A2_EMAIL
_cContacto:=PadR(SA2->A2_CONTATO,30)

mCorpo := 'Sr.(a): ' + _cContacto +Chr(13)+Chr(10)+Chr(13)+Chr(10)
mCorpo += 'Estimado proveedor sigue anexo el Pedido de Compras: ' + cNumPed + ", favor de surtir la mercancia descrita en dicho documento." + Chr(13)+Chr(10)+Chr(13)+Chr(10)
mCorpo += Alltrim(SM0->M0_NOMECOM)+Chr(13)+Chr(10)
mCorpo += 'Teléfono:'+Alltrim(SM0->M0_TEL)+Chr(13)+Chr(10)
mCorpo += 'Site: www.totvs.com.br'+Chr(13)+Chr(10)

Define msDialog oMailPC Title "Pedido de Compras por Email" From 127,037 To 531,774 Pixel 
@ 013,006 To 053,357 Title OemToAnsi("  Datos del Pedido ") 

@ 020,010 Say "Pedido:" Color CLR_HBLUE OF oMailPC Pixel 
@ 020,040 Get cNumPed Picture "@!" When .f. OF oMailPC Pixel Size 40,08
@ 020,097 Say "Email:" Color CLR_HBLUE OF oMailPC Pixel 
@ 020,125 Get _cPara Picture "@" OF oMailPC Pixel Size 150,08 
@ 030,010 Say "Proveedor: " Color CLR_HBLUE OF oMailPC Pixel
@ 030,042 Say SA2->A2_NOME Color CLR_HRED Object oCliente 
@ 040,010 Say "Contacto: " Color CLR_HBLUE OF oMailPC Pixel
@ 040,042 Say _cContacto Object oAutor 
@ 040,042 Get _cContacto Picture "@" OF oMailPC Pixel Size 150,08

@ 80,010 To 182,360 OF oMailPC Pixel
@ 88,015 Get mCorpo MEMO OF oMailPC Pixel Size 340,90

Activate MsDialog oMailPC On Init EnchoiceBar(oMailPC,bOk,bCancel,,) Centered

If nOpc == 1
	cAnexo := 'C:\SPOOL\PedidosCompra\'+cFilename
		EnvMail(cAnexo, cNumPed, _cPara, _cContacto, mCorpo)
EndIf
Return .T.
//-----------------------------------------------------------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------------------------------------------------------

Static Function EnvMail(cAnexo,cNumPed,cPara,cContato,mCorpo)
Private cAssunto     := 'Envio de Pedido de Compras ' + cNumPed
Private nLineSize    := 60
Private nTabSize     := 3
Private lWrap        := .T. 
Private nLine        := 0
Private cTexto       := ""
Private lServErro	 := .T.
Private cServer  	 := Trim(GetMV("MV_RELSERV")) // smtp.tecnotron.ind.br
Private cDe 		 := Trim(GetMV("MV_RELACNT"))
Private cPass    	 := Trim(GetMV("MV_RELPSW"))
Private lAutentic	 := GetMv("MV_RELAUTH",,.F.)
Private aTarget  	 :={cAnexo}
Private nTarget 	 := 0
Private lCheck1 	 := .F.
Private lCheck2 	 := .f.

cCC := UsrRetMail(RetCodUsr())
cAnexos:=cAnexo
CPYT2S(cAnexos,GetSrvProfString("Startpath", ""),.T.)
cAnexos:=GetSrvProfString("Startpath", "")+SubStr(AllTrim(cAnexos),RAT('\',AllTrim(cAnexos))+1)
//msgalert(cAnexo + chr(10) + cAnexos)
lServERRO 	:= .F.
                    
CONNECT SMTP                         ;
SERVER	GetMV("MV_RELSERV"); 	// Nome do servidor de e-mail
ACCOUNT GetMV("MV_RELACNT"); 	// Nome da conta a ser usada no e-mail
PASSWORD GetMV("MV_RELPSW"); 	// Senha
Result lConectou    

lRet := .f.
lEnviado := .f.

If lAutentic
	lRet := Mailauth(cDe,cPass)
Endif
If lRet  
	cPara   := Rtrim(cPara)
	cCC		:= Rtrim(cCC)
	cAssunto:= Rtrim(cAssunto)
	
//	    	    BCC 'diretoria@metalacre.com.br;gerencia@metalacre.com.br';

	SEND MAIL 	FROM cDe ;
		 		To cPara ;
	    	    CC cCc;
		 		SUBJECT	cAssunto;
		 		Body mCorpo;
		 		ATTACHMENT cAnexos;
		 		RESULT lEnviado

	DISCONNECT SMTP SERVER
Endif
If !lConectou .Or. !lEnviado
	cMensagem := ""
	GET MAIL ERROR cMensagem 
	Alert(cMensagem)
else
	ApMsgInfo("Pedido de Compra enviado correctamente", "Envío de Pedido")
Endif
FERASE(cAnexos)
Return