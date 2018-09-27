package com.sage.live.migrationservice.converter;

//this is an issue that account list depends on tempalte and thus on version
//there are a lot of account that I can not create
//TODO FLUSH IN excel for memory?

import java.io.BufferedReader;
//import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;

import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

 class AccountHelper {
    //suffix : "_Classe_FR"
	private static	Integer existingClassesArray[] = {101,102,104,105,106,107,108,109,110,119,120,129,131,138,139,142,143,144,145,146,147,148,151,153,154,155,156,157,158,161,162,163,164,165,166,167,168,169,171,174,178,181,186,187,188,
				201,203,205,206,207,208,211,212,213,214,215,218,231,232,237,238,261,266,267,268,269,271,272,273,274,275,276,277,278,279,280,281,282,290,291,292,293,296,297,311,312,317,321,322,326,
				331,335,341,345,351,355,358,371,372,391,392,393,394,395,397,400,401,403,404,405,408,409,410,411,413,416,418,419,421,422,424,425,426,427,428,431,437,438,441,442,443,444,445,446,447,
				448,449,451,455,456,457,458,462,464,465,467,468,471,472,473,474,475,476,477,478,481,486,487,488,491,495,496,501,502,503,504,505,506,507,508,509,511,512,514,515,516,517,518,519,531,
				532,533,540,580,590,601,602,603,604,605,606,607,608,609,611,612,613,614,615,616,617,618,619,621,622,623,624,625,626,627,628,629,631,633,635,637,641,644,645,646,647,648,651,653,654,
				655,658,661,664,665,666,667,668,671,672,674,675,678,681,686,687,691,695,696,698,699,701,702,703,704,705,706,707,708,709,711,713,721,722,740,751,752,753,754,755,758,761,762,763,764,
				765,766,767,768,771,772,774,775,777,778,781,786,787,791,796,797,890};
		
	//suffix : "_Sous-classe_FR"
	private static	Integer existingSubClassesArray[] = {1061,1062,1063,1064,1068,1671,1674,1675,1681,1685,1687,2801,2803,2805,2807,2808,2810,2811,2812,2813,2814,2815,2818,2900,2905,2906,2907,2908,2930,2931,2932,2966,
				2967,2968,2971,2972,2973,2974,2975,2976,4081,4084,4088,4091,4096,4191,4196,4197,4198,4282,4284,4286,4287,4561,4562,4563,4564,4566,4567,4686,4687,5181,5188,5190,
				6031,6032,6037,6091,6092,6094,6095,6096,6097,6098,6811,6812,6815,6816,6817,7091,7092,7094,7095,7096,7097,7098,16881,16883,16884,16885,16886};
		
	//suffix : "_FR"
	private static	Integer existingAccountsArray[] = {44400,101000,101100,101200,101300,101800,102000,104000,104100,104200,104300,104400,104500,105000,105100,105200,105300,105500,105700,105800,106000,106100,106200,106300,106400,106800,107000,
				108000,109000,110000,119000,120000,129000,131000,131100,131200,131300,131400,131500,131600,131700,131800,138000,139100,139800,142000,142300,142400,143000,143100,143200,144000,145000,146000,
				147000,148000,151000,151100,151200,151300,151400,151500,151600,151800,153000,154000,155000,156000,157000,157200,158000,158100,161000,162000,163000,164000,165000,165100,165500,166000,166100,
				166200,167000,167100,167400,167500,168000,168100,168500,168700,168800,169000,171000,174000,178000,178100,178800,181000,186000,187000,188000,201000,201100,201200,201300,203000,205000,206000,
				207000,208000,208100,211000,211100,211200,211300,211400,211500,211600,212000,213000,213100,213500,213800,214000,215000,215100,215300,215400,215500,215700,218000,218090,218100,218200,218300,
				218400,218500,218600,218700,231000,231200,231300,231500,231800,232000,237000,238000,238200,238300,238500,238800,261000,261100,261800,266000,266100,267000,267100,267400,267500,267600,267700,
				267800,268000,268100,268800,269000,271000,271100,271800,272000,272100,272200,273000,274000,274100,274200,274300,274800,275000,275100,275500,276000,276100,276800,277100,277200,278000,279000,
				280000,280100,280300,280500,280700,280800,281000,281100,281200,281300,281400,281500,281800,282000,290000,290500,290600,290700,290800,291000,292000,293000,293100,293200,296000,296100,296600,
				296700,296800,297000,297100,297200,297300,297400,297500,297600,311000,312000,317000,321000,321100,321200,322000,322100,322200,322300,322400,322500,326000,326100,326500,326700,331000,331100,
				331200,335000,335100,335200,341000,341100,341200,345000,345100,345200,351000,351100,351200,355000,355100,355200,358000,358100,358500,358600,371000,372000,391000,391100,391200,391700,392000,
				392100,392200,392600,393000,393100,393500,394000,394100,394500,395000,395100,395500,397000,397100,397200,400000,401000,401100,401700,403000,404000,404100,404700,405000,408000,408100,408400,
				408800,409000,409100,409600,409700,409800,410000,411000,411100,411700,413000,416000,418000,418100,418800,419000,419100,419600,419700,419800,421000,422000,424000,424600,424800,425000,426000,
				427000,428000,428200,428400,428600,428700,431000,437000,437100,437200,437300,437400,437500,437600,438000,438200,438600,438700,441000,441100,441700,441800,441900,442000,442400,442500,443000,
				443100,443800,445000,445200,445500,445600,445700,445800,446000,447000,448000,448200,448600,448700,449000,451000,455000,455100,455800,456000,456100,456200,456300,456400,456600,456700,457000,
				458000,458100,458800,462000,464000,465000,467000,468000,468600,468700,471000,472000,473000,474000,475000,476000,476100,476200,476800,477000,477100,477200,477800,478000,481000,481600,486000,
				487000,488000,488600,488700,491000,495000,495100,495500,495800,496000,496200,496500,496700,501000,502000,502100,502200,503000,503100,503500,504000,505000,506000,506100,506500,507000,508000,
				508100,508200,508800,509000,511000,511100,511200,511300,511400,511500,512000,512010,512100,512400,514000,515000,516000,517000,518000,518100,518800,519000,519100,519300,519800,531000,531100,
				531400,532000,533000,540000,580000,590000,590300,590400,590600,590800,601000,601090,601100,601200,601700,601900,602000,602100,602200,602600,602900,603000,603100,603200,603700,604000,604010,
				604020,604021,604055,604085,604090,604900,605000,605900,606000,606100,606300,606400,606800,606900,607000,607010,607020,607021,607055,607085,607100,607200,607900,608000,608900,609000,609100,
				609200,609400,609500,609600,609700,609800,611000,612000,612200,612500,613000,613200,613500,613600,614000,615000,615200,615500,615600,616000,616100,616200,616300,616400,616500,617000,618000,
				618100,618300,618500,619000,621000,621100,621400,622000,622100,622200,622400,622500,622600,622700,622800,623000,623100,623200,623300,623400,623500,623600,623700,623800,624000,624100,624200,
				624300,624400,624700,624800,625000,625100,625500,625600,625700,626000,627000,627100,627200,627500,627600,627800,628000,628100,628400,629000,631100,631200,631300,631400,631800,633000,633100,
				633200,633300,633400,633500,633800,635000,635100,635200,635300,635400,635800,637000,637100,637200,637400,637800,641000,641100,641200,641210,641300,641400,641500,641600,644000,645000,645100,
				645200,645300,645310,645320,645400,645500,645800,645810,645820,646000,647000,647100,647200,647300,647400,647500,648000,651000,651100,651600,651800,653000,654000,654100,654400,655000,655100,
				655500,658000,661000,661100,661200,661500,661600,661700,661800,664000,665000,666000,667000,668000,671000,671100,671200,671300,671400,671500,671700,671800,672000,674000,674100,674200,675000,
				675100,675200,675600,675800,678000,678100,678200,678300,678800,681000,681100,681200,681500,681600,681700,686000,686100,686500,686600,686800,687000,687100,687200,687300,687400,687500,687600,
				691000,695000,695100,695200,695400,696000,698000,698100,698900,699000,701000,701010,701020,701021,701055,701085,701090,701100,701120,701130,701140,701150,701160,701170,701200,702000,702120,
				702130,702160,703000,703120,703130,703160,704000,704100,704120,704130,704160,704200,705000,705120,705130,705160,706000,706010,706020,706021,706055,706085,706090,706120,706130,706150,706160,
				706170,707000,707010,707020,707021,707055,707085,707090,707100,707120,707130,707150,707160,707170,707200,708000,708100,708120,708130,708160,708200,708300,708400,708500,708600,708700,708800,
				709000,709100,709200,709400,709500,709600,709700,709800,711000,713000,713300,713400,713500,721000,722000,740000,751000,751100,751600,751800,752000,753000,754000,755000,755100,755500,758000,
				761000,761100,761200,761600,761700,762000,762100,762600,762700,763000,763100,763800,764000,765000,766000,767000,768000,771000,771100,771300,771400,771500,771700,771800,772000,774000,774100,
				774200,775000,775100,775200,775600,775800,777000,778000,778100,778200,778300,778800,781000,781100,781500,781600,781700,786000,786500,786600,787000,787200,787300,787400,787500,787600,791000,
				796000,797000,890000,1391100,1391200,1391300,1391400,1391500,1391600,1391700,1391800,1688100,1688300,1688400,1688500,1688600,1688700,1688800,2012100,2012200,2115100,2115500,2115800,2131100,
				2131500,2131800,2135100,2135500,2135800,2138100,2138200,2138300,2138400,2138500,2151100,2151400,2153100,2153400,2768200,2768400,2768500,2768800,2808100,2818700,2908100,2918700,2978700,4097100,
				4097400,4452920,4455100,4455800,4456200,4456210,4456220,4456221,4456255,4456300,4456600,4456605,4456610,4456620,4456621,4456685,4456700,4456800,4456900,4456920,4457020,4457100,4457102,4457109,
				4457110,4457113,4457120,4457121,4457125,4457126,4457127,4457128,4457155,4457185,4457800,4458100,4458200,4458300,4458400,4458600,4458700,4458810,4458820,4561100,4561500,4562100,4562500,6021100,
				6021200,6022100,6022200,6022300,6022400,6022500,6026100,6026500,6026700,6040211,6070211,6163600,6163700,6163800,6351100,6351200,6351300,6351400,6354100,6611600,6611700,6618100,6618800,6811100,
				6811200,6816100,6816200,6817300,6817400,6866200,6866500,6872500,7010211,7010212,7010213,7010214,7010215,7010216,7010217,7060211,7060212,7060213,7060214,7060215,7060216,7060217,7070211,7070212,
				7070213,7070214,7070215,7070216,7070217,7133100,7133500,7134100,7134500,7135100,7135500,7135800,7811100,7811200,7816100,7816200,7817300,7817400,7866200,7866500,7872500,7872600,7872700,21158100,
				21158800,21318100,21318800,21358100,21358800,44566211,44571091,44571131,44571251,44571261,44571271,44571281,445200201,445200202,445200203,445200204,445200212,445201001,445201002,445201003,
				445201004,445202001,445202002,445202003,445202004,445205501,445205502,445205503,445205504,445208501,445208502,445208503,445208504,445710201,445710202,445710203,445710204,445710205,445710206,
				445710210,445710221,445710222,445710223,445710224,445711001,445711002,445711003,445711004,445711005,445711010,445711020,445712001,445712002,445712003,445712004,445712005,445712010,445715501,
				445715502,445715503,445715504,445715505,445715510,445715520,445718501,445718502,445718503,445718504,445718505,445718510};
	
	private static String InternalFind(String aAccount, Integer aLen, Integer[] aArr) {
		String tmp;
		if (aLen>0) {
			if (aAccount.length()>= aLen)
				tmp = aAccount.substring(0, aLen);
			else
				return "";
		}
		else 
			tmp = aAccount;
		
		Integer	thisAccount = Integer.parseInt(tmp);
		if (Arrays.binarySearch(aArr,thisAccount)>=0)
			return thisAccount.toString();
		else 
			return "";		
	}		
	
	public static String getFinancialIdentifier(String aAccount) {
		//aAccount is expected to be numerical only
		String res = InternalFind(aAccount,-1,existingAccountsArray);
		if (!res.isEmpty())
			res += "_FR";
		return res;
	}
	
	public static String getSubClasse(String aAccount) {
		//aAccount is expected to be numerical only
		String res = InternalFind(aAccount,5,existingSubClassesArray);
		if (res.isEmpty())
			res = InternalFind(aAccount,4,existingSubClassesArray);		
		
		if (!res.isEmpty())
			res += "_Sous-classe_FR";
		
		return res;
	}
	
	public static String getClasse(String aAccount) {
		//aAccount is expected to be numerical only
		String res = InternalFind(aAccount,3,existingClassesArray);
		
		if (!res.isEmpty())
			res += "_Classe_FR";
		
		return res;
	}

 }

class FECPArseException extends Exception{ 
	 
	 public FECPArseException(String message) {
		    super(message);
		}
	 static final long serialVersionUID = 1;
 }

 class FECLine {
	
	public String journalCode;//1. Code du journal de l��criture comptable
	public String journalCaption;//2. Libell� du journal de l��criture comptable
	public String ledgerID;//3. Num�ro de l��criture comptable (devant �tre bas� sur une s�quence continue)	
	public Date comptablizationDate;//4. Date de comptabilisation de l��criture comptable
	public String ledgerAccountCode;//5. Num�ro de compte (compte r�ellement utilis�e par l�entreprise, en fonction de son plan comptable)
	public String ledgerAccountCaption;//6. Libell� de compte
	public String auxAccountCode;//7. Num�ro de compte auxiliaire (� blanc si non utilis�)
	public String auxAccountCaption;//8. Libell� de compte auxiliaire (� blanc si non utilis�)
	public String reference;//9. R�f�rence de la pi�ce justificative
	public Date pieceDate;//10. Date de la pi�ce justificative
	public String caption;//11. Libell� de l��criture comptable (identification litt�rale du motif de l��criture comptable)
	public Number debtAmount;//12. Montant au d�bit
	public Number credAmount;//13. Montant au cr�dit
	public String lettrage;//14. Lettrage de l��criture comptable (� blanc si non utilis�)
	public Date lettrageDate;//15. Date de lettrage (� blanc si non utilis�)
	public Date validationDate;//16. Date de validation de l��criture comptable
	public Number amount;//17. Montant en devise (� blanc si non utkilis�)
	public String currency;//18. Identifiant de la devise (� blanc si non utilis�)	
	public FECLine(String aLine, String aSeparator) throws FECPArseException{
		
		//TODO others colums :  CodeEtbt	TypePiece	eDate	RefOrigine	Numero	NumeroDef	OrigineLot

		DateFormat format = new SimpleDateFormat("yyyyMMdd", Locale.FRANCE);
		NumberFormat formatter = NumberFormat.getInstance(Locale.FRANCE) ;

		
		String[] tokens;
		tokens = aLine.split(aSeparator);
		if (tokens.length>=18){			
			journalCode = tokens[0];
			journalCaption= tokens[1];
			ledgerID= tokens[2];
			if (tokens[3]!=""){
				try{
					comptablizationDate= format.parse(tokens[3]);
				} catch (ParseException ex) {
					throw new FECPArseException("can not parse comptablization Date "+tokens[3]);					
				}
			}
			ledgerAccountCode= tokens[4];
			ledgerAccountCaption= tokens[5];
			auxAccountCode= tokens[6];
			auxAccountCaption= tokens[7];
			reference= tokens[8];
			if (!tokens[9].isEmpty()){
				try{
					pieceDate = format.parse(tokens[9]);
				} catch (ParseException ex) {
					throw new FECPArseException("can not parse piece Date "+tokens[9]);					
				}
			}
			caption= tokens[10];			
			try{
				debtAmount = formatter.parse(tokens[11]);
			} catch (ArithmeticException | ParseException ex) {
				
				throw new FECPArseException("can not parse debt amount "+tokens[11]);					
			}			
			try{
				credAmount = formatter.parse(tokens[12]);
			} catch (ArithmeticException | ParseException ex) {
				
				throw new FECPArseException("can not parse cred amount "+tokens[12]);					
			}
			lettrage= tokens[13];
			if (!tokens[14].isEmpty() ){
				try{
					lettrageDate= format.parse(tokens[14]);
				} catch (ParseException ex) {
					
					throw new FECPArseException("can not parse lettrage Date "+tokens[14]);					
				}
			}
			
			if (!tokens[15].isEmpty()){
				try{
					comptablizationDate= format.parse(tokens[15]);
				} catch (ParseException ex) {
					throw new FECPArseException("can not parse validation Date "+tokens[15]);					
				}
			}
			
			//montant en devise
			if (!tokens[16].isEmpty()){
				try{
					amount = formatter.parse(tokens[16]);
				} catch (ArithmeticException | ParseException ex) {
					throw new FECPArseException("can not parse cred amount "+tokens[16]);					
				}
			}			
			
			currency= tokens[17];
			
		}
		else
			throw new FECPArseException("invalid FEC line "+aLine);
	}
	
} 
 
abstract class Sheet{
	
	static final Integer firstDataLine = 2;//0 indexed
	
	protected SXSSFSheet _XSSSheet;
	protected ConfSheet _ConfSheet;
	
	public Sheet(ConfSheet aConfSheet) {
		_XSSSheet = aConfSheet.getWorkBook().createSheet(getName());
		aConfSheet.AddSheet(this);
		_ConfSheet = aConfSheet;
		Init();
	}
	
	public Sheet() {
	}
	
	protected abstract  String getName();	
	protected abstract void Internalsave(Integer aLineNumber);
	protected abstract void Init() ;
	
	protected Integer getLineNumber(Boolean aHeader) {
		if (aHeader) {
			return 0;
		} else {
			Integer ll = _XSSSheet.getLastRowNum() +1;
			if (ll<firstDataLine)
				return firstDataLine;
			else
				return ll;
		}
	}
	
	protected void doWriteLine(Integer lineNumber,Object... cols) {
		SXSSFRow res = _XSSSheet.createRow(lineNumber);
		Integer cpt = 0; 
		for (Object o:cols) {
			
			if (o instanceof String) {
				res.createCell(cpt++).setCellValue((String)o);
			} else if (o instanceof Double) {
					res.createCell(cpt++).setCellValue((Double)o);
				} else if (o instanceof Integer) {
						res.createCell(cpt++).setCellValue((Integer)o);
				} else {
					assert false;
			}			
		}
	}
	
	public void InitColumns(Object... cols) {
		doWriteLine(getLineNumber(true) ,cols);
	}
	
	protected void writeLine(Object... cols) {
		doWriteLine(getLineNumber(false) ,cols);
	}
	
	public SXSSFSheet getXSSSheet(){
		return _XSSSheet;
	}
	
	
}

class CustomersSheet extends Sheet {
	
	public CustomersSheet(ConfSheet aConfSheet) {
		super(aConfSheet);
	}

	protected String getName()  {
		return "Customers";
	}
	
	protected  void Init()  {
		InitColumns("Account.Name",
				 "Account.AccountNumber",
				 "Account.s2cor__Is_Active__c",
				 "Account.s2cor__Sage_UID__c");
	}
	
	protected void Internalsave(Integer aLineNumber) {		
		Integer fl = firstDataLine+1;
		Integer ll = getLineNumber(false);
		_ConfSheet.doWriteLine(aLineNumber,aLineNumber,getName(),"SObjectAndDimension","Account","Customer%","s2cor__Account__c","Name","","",fl,ll,"");
	}
	
}

class SuppliersSheet  extends Sheet {
	
	public SuppliersSheet(ConfSheet aConfSheet) {
		super(aConfSheet);		
	}

	protected String getName()  {
		return "Suppliers";
	}
	
	protected  void Init()  {
		 InitColumns("Account.Name",
				 "Account.AccountNumber",
				 "Account.s2cor__Is_Active__c",
				 "Account.s2cor__Sage_UID__c");		
	}
	
	protected void Internalsave(Integer aLineNumber) {
		Integer fl = firstDataLine+1;
		Integer ll = getLineNumber(false);
		_ConfSheet.doWriteLine(aLineNumber,aLineNumber,getName(),"SObjectAndDimension","Account","Supplier%","s2cor__Account__c","Name","","",fl,ll);
	}
	
	
}

class ChartOfAccountSheet  extends Sheet {
	
	public ChartOfAccountSheet(ConfSheet aConfSheet) {
		super(aConfSheet);		
	}

	protected String getName()  {
		return "Chart of accounts";
	}
	
	protected  void Init()  {
		InitColumns(	"CustomerLedgerAccountName",
				"s2cor__Sage_ACC_Ledger_Account__c.s2cor__UID__c",
				"Active",
				"RenameAccount",
				"s2cor__Sage_ACC_Ledger_Account__c.s2cor__Parent__c.s2cor__UID__c",
				"s2cor__Sage_ACC_Dimension__c.Name",
				"Debit",
				"Credit");
	}
	
	protected void Internalsave(Integer aLineNumber) {
		Integer fl = firstDataLine+1;
		Integer ll = getLineNumber(false);
		DateFormat ff = new SimpleDateFormat("yyyy-MM-dd", Locale.FRANCE);
		String dd = ff.format(Calendar.getInstance().getTime());//This date is used as the ledger entry date for the opening balances for the ledger accounts.
		_ConfSheet.doWriteLine(aLineNumber,aLineNumber,getName(),"LedgerAccounts","s2cor__Sage_ACC_Ledger_Account__c","","","","","",fl,ll,dd);
	}
	
	
}

class LedgerHistorySheet extends Sheet {
	
	public LedgerHistorySheet(ConfSheet aConfSheet) {
		super(aConfSheet);		
	}

	protected String getName()  {
		return "LedgerHistory";
	}
	
	protected  void Init()  {
		InitColumns("s2cor__Sage_ACC_Ledger_Entry__c.s2cor__UID__c",	
				"s2cor__Sage_ACC_Ledger_Item__c.s2cor__UID__c",
				"s2cor__Sage_ACC_Ledger_Entry__c.s2cor__Code__c",
				"s2cor__Sage_ACC_Ledger_Entry__c.s2cor__Date__c",
				"s2cor__Sage_ACC_Ledger_Item__c.s2cor__Description__c",
				"s2cor__Sage_ACC_Ledger_Item__c.s2cor__Ledger_Account__r.s2cor__UID__c",
				"s2cor__Sage_ACC_Ledger_Item__c.s2cor__Base_Debit__c",
				"s2cor__Sage_ACC_Ledger_Item__c.s2cor__Base_Credit__c",
				"s2cor__Sage_ACC_Ledger_Item__c.s2cor__Foreign_Debit__c",
				"s2cor__Sage_ACC_Ledger_Item__c.s2cor__Foreign_Credit__c",
				"s2cor__Sage_ACC_ledger_Entry_Tax_Breakdown__c.s2cor__Tax_Amount__c",
				"s2cor__Sage_ACC_ledger_Entry_Tax_Breakdown__c.s2cor__Tax_Rate__c.s2cor__UID__c",
				"s2cor__Sage_ACC_Ledger_Item__c.s2cor__Currency__c.s2cor__UID__c",
				"s2cor__Sage_ACC_Ledger_Item__c.s2cor__Exchange_Rate__c",
				"[Customer%]",
				"[Supplier%]",
				"[SourceClassification%]",
				"[Code_Lettrage%]");
	}
	
	protected void Internalsave(Integer aLineNumber) {
		Integer fl = firstDataLine+1;
		Integer ll = getLineNumber(false);
		_ConfSheet.doWriteLine(aLineNumber,aLineNumber,getName(),"LedgerHistory","s2cor__Sage_ACC_Ledger_Entry__c","","","","","",fl,ll);
	}
	
}

class JournalsDimensionsSheet  extends Sheet {
	
	public JournalsDimensionsSheet(ConfSheet aConfSheet) {
		super(aConfSheet);		
	}

	protected String getName()  {
		return "Journals Dimensions";
	}
	
	protected  void Init()  {
		InitColumns("[SourceClassification%].Name",	
				  "s2cor__Code__c");
	}
	
	protected void Internalsave(Integer aLineNumber) {
		Integer fl = firstDataLine+1;
		Integer ll = getLineNumber(false);
		_ConfSheet.doWriteLine(aLineNumber,aLineNumber,getName(),"Dimension","s2cor__Sage_ACC_Tag__c","","","","","",fl,ll);
	}
	
}

class CustomersDimensionsSheet  extends Sheet {
	
	public CustomersDimensionsSheet(ConfSheet aConfSheet) {
		super(aConfSheet);		
	}

	protected String getName()  {
		return "Customers Dimensions";
	}
	
	protected  void Init()  {
		InitColumns("[Customer%].Name",
					"s2cor__Code__c"); 
	}
	
	protected void Internalsave(Integer aLineNumber) {
		Integer fl = firstDataLine+1;
		Integer ll = getLineNumber(false);
		_ConfSheet.doWriteLine(aLineNumber,aLineNumber,getName(),"Dimension","s2cor__Sage_ACC_Tag__c","","","","","",fl,ll);
	}
	
}

class SuppliersDimensionsSheet extends Sheet {
	
	public SuppliersDimensionsSheet(ConfSheet aConfSheet) {
		super(aConfSheet);		
	}

	protected String getName()  {
		return "Suppliers Dimensions";
	}
	
	protected  void Init()  {
		InitColumns("[Supplier%].Name",
				   "s2cor__Code__c"); 
	}
	
	protected void Internalsave(Integer aLineNumber) {
		Integer fl = firstDataLine+1;
		Integer ll = getLineNumber(false);
		_ConfSheet.doWriteLine(aLineNumber,aLineNumber,getName(),"Dimension","s2cor__Sage_ACC_Tag__c","","","","","",fl,ll);
	}
	
}

class LettragesSheet  extends Sheet {
	
	public LettragesSheet(ConfSheet aConfSheet) {
		super(aConfSheet);		
	}

	protected String getName()  {
		return "Lettrages";
	}
	
	protected  void Init()  {
		InitColumns("[Code_Lettrage%].Name",
				 "s2cor__Reconciled_Date__c");	
	}
	
	protected void Internalsave(Integer aLineNumber) {
		Integer fl = firstDataLine+1;
		Integer ll = getLineNumber(false);
		_ConfSheet.doWriteLine(aLineNumber,aLineNumber,getName(),"Dimension","s2cor__Sage_ACC_Tag__c","","","","","",fl,ll);
	}
	
}


class ConfSheet extends Sheet {

	private List<Sheet>_sheets;
	private SXSSFWorkbook _Workbook;

	public ConfSheet() {
		_sheets = new ArrayList<Sheet>();
		_Workbook = new SXSSFWorkbook(100);	
		_XSSSheet = _Workbook.createSheet(getName());
		Init();
	}

	protected void Internalsave(Integer aLineNumber) {
		
	}	
	
	protected String getName()  {
		return "Configuration";
	}
	
	
	protected void Init(){
		InitColumns("Order",
				"Sheet",
				"Type",
				"Main Sobject",
				"Main Dimension Type",
				"Dimension Link",
				"Tag Value",
				"Account",
				"Contraaccount",
				"Data Start Row",
				"Data End Row","Date");
	}
	
	public SXSSFWorkbook getWorkBook(){
		return _Workbook; 
	}
	
	public void AddSheet(Sheet aSheet) {
		_sheets.add(aSheet);
	}
	
	protected void save(String aoutfn) throws IOException {
		Integer cpt = 1;
		for (Sheet s: _sheets) {			
			s.Internalsave(cpt);
			cpt++;
		}
		
		FileOutputStream outputStream = new FileOutputStream(aoutfn);
		_Workbook.write(outputStream);		
      
	}
	
	protected void finalize ()  {
		_Workbook.dispose();
		
	}
}

class Account{	

	public Boolean haveToCreate = false;
	public String originAccount = "";
	public String originAccountCaption = "";
	public String CurrentLedgerAccount = "";//this is a caption
	public String FinancialsIdentifier = "";//UID
	public String ParentUID = "";//only for creation
	//private String Legislation;
	public String auxAccountCode;//. Num�ro de compte auxiliaire (� blanc si non utilis�)
	public String auxAccountCaption;//8. Libell� de compte auxiliaire (� blanc si non utilis�)
		
	
	public Account(FECLine aFECLine) throws FECPArseException {
		
		originAccount = aFECLine.ledgerAccountCode;//5. Num�ro de compte (compte r�ellement utilis�e par l�entreprise, en fonction de son plan comptable)
		originAccountCaption = aFECLine.ledgerAccountCaption;//6. Libell� de compte
		CurrentLedgerAccount = originAccount + " " + originAccountCaption;
		FinancialsIdentifier = "";
		auxAccountCode = aFECLine.auxAccountCode;
		auxAccountCaption = aFECLine.auxAccountCaption;
		
		
		String tmp = originAccount;
		char[] chArr = originAccount.toCharArray();
		for (int i = 0 ; i<chArr.length; i++){
			if (! Character.isDigit(chArr[i])){
				tmp = originAccount.substring(0,i);
				break;
			}
		}		
		
		if (tmp.length()>=3){			
			//FinancialsIdentifier = AccountHelper.getFinancialIdentifier(tmp);
			//haveToCreate = FinancialsIdentifier.isEmpty();		 
			haveToCreate = true;
			
			if (FinancialsIdentifier.isEmpty()){
				ParentUID = AccountHelper.getSubClasse(tmp);
				if (!ParentUID.isEmpty()) {
					FinancialsIdentifier = originAccount + "_FR";
					haveToCreate = true;					
				}				
			}
			
			if (FinancialsIdentifier.isEmpty()){
				ParentUID = AccountHelper.getClasse(tmp);
				if (!ParentUID.isEmpty()) {
					FinancialsIdentifier = originAccount + "_FR";
					haveToCreate = true;
				}							
			}			
			
			if (FinancialsIdentifier.isEmpty()){
				throw new FECPArseException("invalid account "+originAccount);					
			}
			
		} else {
			throw new FECPArseException("invalid account "+originAccount);					
		}
	}
	
	public Boolean IsSupplier() {
		return FinancialsIdentifier.startsWith("40");
	}
	
	public Boolean IsCustomer() {
		return FinancialsIdentifier.startsWith("41");		
	}
		

	
	public void AddRow(ChartOfAccountSheet aSheet){
		//account
		if (haveToCreate){
			aSheet.writeLine(CurrentLedgerAccount,
							 FinancialsIdentifier,
							 "TRUE",
							 "",
							 ParentUID,							 
							 "",
							 "",
							 "");
		}
	}
	
	private void AddThirdPartyRow(Sheet aSheet){		
		aSheet.writeLine(CurrentLedgerAccount,originAccount,"TRUE",FinancialsIdentifier);
	}
	
	public void AddSupplierRow(Sheet aSheet){
		if (FinancialsIdentifier.startsWith("40")){
			AddThirdPartyRow(aSheet);
		}		
	}
	
	public void AddCustomerRow(Sheet aSheet){
		if (FinancialsIdentifier.startsWith("41")){
			AddThirdPartyRow(aSheet);
		}		
	}

}

class Ledger implements Comparable<Ledger>{
	
	public String reference = "";
	public Integer lineReference = 0;
	private String ledgerReference = "";
	private Date ledgerDate;
	private Account account = null;
	private String Caption = "";
	private Number debit = 0;
	private Number credit = 0;
	private Number foreignDebit = 0;
	private Number foreignCredit = 0;	
	private String currency = "";
	public String journal = "";
	public String journalCaption = "";
	public String lettrage = "";
	public Date lettrageDate;
	
	private String auxAccountCode = "";
	private String auxAccountCaption = "";
	public Ledger(FECLine aFECLine, Account aAccount){
		
		account = aAccount;
		journal = aFECLine.journalCode;
		journalCaption = aFECLine.journalCaption;//2. Libell� du journal de l��criture comptable
		reference = aFECLine.ledgerID;
		ledgerDate = aFECLine.comptablizationDate;

		auxAccountCode = aFECLine.auxAccountCode;//7. Num�ro de compte auxiliaire (� blanc si non utilis�)
		auxAccountCaption = aFECLine.auxAccountCaption;//8. Libell� de compte auxiliaire (� blanc si non utilis�)
		
		ledgerReference = aFECLine.reference;
		//aFECLine.pieceDate;//10. Date de la pi�ce justificative
		Caption = aFECLine.caption;

		debit = aFECLine.debtAmount;
		credit = aFECLine.credAmount;
		lettrage = aFECLine.lettrage;
		lettrageDate = aFECLine.lettrageDate;
		//aFECLine.validationDate;//16. Date de validation de l��criture comptable
		if ((aFECLine.amount != null) && (aFECLine.amount.intValue()!=0)){
			if (debit.intValue()!=0)
				foreignDebit = aFECLine.amount;
			else
				foreignCredit = aFECLine.amount;
		}		
		currency = aFECLine.currency;//18. Identifiant de la devise (� blanc si non utilis�)	
		if (currency.isEmpty()){
			currency = "EUR";
		}
		
	}
	
	
	
	public void AddRow(LedgerHistorySheet aSheet){
		
		DateFormat df = new SimpleDateFormat("yyyy-MM-dd", Locale.FRANCE);	
		
		double exRate = 0;
		if (debit.doubleValue()+credit.doubleValue()!=0)
			exRate = (foreignCredit.doubleValue()+foreignDebit.doubleValue())/(debit.doubleValue()+credit.doubleValue());
		
		String auxAccountCaptionCustomer = "";
		String auxAccountCaptionSupplier = "";
		
		if (auxAccountCode.startsWith("41")){
			auxAccountCaptionSupplier = auxAccountCaption;
		}		
		
		if (auxAccountCode.startsWith("40")){
			auxAccountCaptionCustomer = auxAccountCaption;
		}		
		
		
		aSheet.writeLine(	reference,
						 	reference + "-" + lineReference.toString(),
							ledgerReference,
							df.format(ledgerDate),
							Caption,
							account.FinancialsIdentifier,
							debit.doubleValue(),
							credit.doubleValue(),		
							foreignDebit.doubleValue(),
							foreignCredit.doubleValue(),
							"",//tax amount		
							"",//tax rate		
							currency,
							exRate,
							auxAccountCaptionSupplier,
							auxAccountCaptionCustomer,
							journalCaption,
							lettrage);
	}
	
	@Override
    public int compareTo(Ledger l) {
		return reference.compareTo(l.reference);
    }
	
}

class Journal{
	private String _code = "";
	private String _caption = "";
	public Journal(String aCode, String aCaption) {
		_code = aCode;
		_caption = aCaption;
	}
	
	
	
	public void AddRow(JournalsDimensionsSheet aSheet){
		aSheet.writeLine(_caption,_code);
	}
	
	public void AddSupplierRow(SuppliersDimensionsSheet aSheet){
		aSheet.writeLine(_caption,_code);
	}
	
	public void AddCustomerRow(CustomersDimensionsSheet aSheet){
		aSheet.writeLine(_caption,_code);
	}
}

class Lettrage{
	private String _codeLettrage = "";
	private Date _dateLettrage;
	public Lettrage(String aCodeLettrage, Date aDateLettrage) {
		_codeLettrage = aCodeLettrage;
		_dateLettrage = aDateLettrage;			
	}	
	
	public void AddRow(LettragesSheet aSheet){
		DateFormat df = new SimpleDateFormat("yyyy-MM-dd", Locale.FRANCE);
		aSheet.writeLine(_codeLettrage,df.format(_dateLettrage));
	}
	
}
	
public class FEC2xlsx {
	
	private List<Account> accountList;
	private List<Ledger> ledgerList;
	private HashMap<String,Journal> journalMap;
	private HashMap<String,Journal> customerJournalMap;
	private HashMap<String,Journal> supplierJournalMap;
	private HashMap<String,Lettrage> lettrageMap;
	
	
	public FEC2xlsx(){		
		accountList = new ArrayList<>();
		ledgerList = new ArrayList<>();	
		journalMap = new HashMap<String,Journal>();	
		customerJournalMap = new HashMap<String,Journal>();	
		supplierJournalMap = new HashMap<String,Journal>();	
		lettrageMap = new HashMap<String,Lettrage>();	
			
	}
	
	private Account getAccount(FECLine aFecLine) throws FECPArseException {
		
		for (int i = 0 ; i< accountList.size(); i++){
			if (accountList.get(i).originAccount.equals(aFecLine.ledgerAccountCode) )
				return accountList.get(i);			
		}
		
		Account acc = new Account(aFecLine);
		accountList.add(acc);
		return acc;
	}	
		
	private void FECLineToInternal(FECLine aFecLine) throws FECPArseException {
		 
		 Account aAccount = getAccount(aFecLine);
		 Ledger l = new Ledger(aFecLine,aAccount);
		 l.lineReference = 1;
		 if (ledgerList.size()>0){
			 Ledger lastLedger = ledgerList.get(ledgerList.size()-1);
			 if (lastLedger.reference.equals(l.reference)){
				 l.lineReference = lastLedger.lineReference+1;
			 } 		 
		 }		 
		 ledgerList.add(l);		 
	}	
	
	
	public boolean Convert(String afn, String aoutfn, IConvertLogger<LogType, String> alogger) throws IOException {	

		//FEC to internal
		Boolean result = true;
		String thisLine;
		String sep = "";
		String[] seps = {"\t",";","|"};
		BufferedReader br = new BufferedReader(new FileReader(afn));
	    try{
			try {
				thisLine = br.readLine();//ignore first line
				for (Integer i = 0 ; i<seps.length ; i++){
					if (thisLine.contains(seps[i])){
						sep = seps[i];
						break;
					}
				}
				
				if (sep=="") {
					throw new FECPArseException("can not parse this file, unknown separator ");
				}
					
				while ((thisLine = br.readLine()) != null) {				
					try {					
						FECLine l = new FECLine(thisLine,sep);
						FECLineToInternal(l);					
						
					} catch (FECPArseException e) {					
						alogger.doLog(LogType.error,"convert : "+e.getMessage());
						result = false;
					}			     
				}
			} catch (IOException | FECPArseException e) {			
				alogger.doLog(LogType.error,"convert : "+e.getMessage());
				result = false;
			}
		} finally {
			br.close();
		}
	    
	    //do stuff on internal
	    Collections.sort(ledgerList);
	    Ledger lastLedger = null;
	    for (Integer i = 0 ; i < ledgerList.size() ; i++){
	    	Ledger l = ledgerList.get(i);
	    	//numerotation
	    	if (lastLedger!=null){
	    		if (lastLedger.reference.equals(l.reference)){
	    			l.lineReference = lastLedger.lineReference+1;
	    		} else {
	    			l.lineReference = 1;
	    		}
	    	} else{
	    		l.lineReference = 1;
	    	}
	    	lastLedger = l;
	    
	    	//journal
	    	if (! l.journal.isEmpty()) {
	    		journalMap.putIfAbsent(l.journal, new Journal(l.journal, l.journalCaption));
	    	}
	    	//lettrage
	    	if (! l.lettrage.isEmpty()) {
	    		lettrageMap.putIfAbsent(l.lettrage, new Lettrage(l.lettrage, l.lettrageDate));
	    	}
	    }
	    
	    for (int i=0;i<accountList.size();i++){
	    	Account acc = accountList.get(i);
	    	if (acc.IsCustomer()) {
	    		if (!acc.auxAccountCaption.isEmpty())
	    			customerJournalMap.put(acc.auxAccountCode, new Journal(acc.auxAccountCode,acc.auxAccountCaption));
	    	} else {
	    		if (acc.IsSupplier()) {
	    			if (!acc.auxAccountCaption.isEmpty())
		    			supplierJournalMap.put(acc.auxAccountCode, new Journal(acc.auxAccountCode,acc.auxAccountCaption));	    			
	    		}
	    	}
	    }	    
	    
	    
	    //Internal to workbook
	    ConfSheet confs = new ConfSheet();
		
		CustomersSheet cs = new CustomersSheet(confs);
		SuppliersSheet ss = new SuppliersSheet(confs);
		ChartOfAccountSheet cas = new ChartOfAccountSheet(confs);
		LedgerHistorySheet lhs = new LedgerHistorySheet(confs);
		JournalsDimensionsSheet jds = new JournalsDimensionsSheet(confs);
		CustomersDimensionsSheet cds = new CustomersDimensionsSheet(confs);
		SuppliersDimensionsSheet sds = new SuppliersDimensionsSheet(confs);
		LettragesSheet ls = new LettragesSheet(confs);
		
		
		for (int i=0;i<accountList.size();i++){
			accountList.get(i).AddRow(cas);
			accountList.get(i).AddSupplierRow(ss);
			accountList.get(i).AddCustomerRow(cs);
		}
		
		for (int i=0;i<ledgerList.size();i++){
			ledgerList.get(i).AddRow(lhs);			
		}
		
		for (Journal j : journalMap.values()){
			j.AddRow(jds);			
		}	
		
		for (Journal j : customerJournalMap.values()){
			j.AddCustomerRow(cds);			
		}		
		
		for (Journal j : supplierJournalMap.values()){
			j.AddSupplierRow(sds);			
		}
		
		for (Lettrage l : lettrageMap.values()){
			l.AddRow(ls);			
		}
		
		try {			
			confs.save(aoutfn);
		} catch (Exception e){
			alogger.doLog(LogType.error,"save : "+e.getMessage());
			result = false;			
		}
		
		return result;
	    
	}
	
	public static void main(String[] args) throws IOException {
		FEC2xlsx x = new FEC2xlsx();
		x.Convert(args[0],args[1],(aLogType,aMessage) -> System.out.println(aLogType.toString() + " " + aMessage));		
    }
}


