package net.sf.jxls.sample.dynamicColumns;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 * This sample was contributed by Senthur Selvi Karunakaran
 */
public class DynamicXLSGeneration {

    private static String templateFileName = "examples/templates/ex_temp.xls";
    private static String destFileName = "build/ex_output.xls";

	@SuppressWarnings("unchecked")
	public static void main(String[] args) throws ParsePropertyException, IOException, InvalidFormatException {
        if (args.length >= 2) {
            templateFileName = args[0];
            destFileName = args[1];
        }

		List<ValueBean> contractList = new ArrayList<ValueBean>();
		
		String afvdat_brk ="";	
		String dlnrnr = "";		
		String echtsnr ="";
		String gebdat_brk = "";
		String gesl_brk ="";
		String nm_dlnr_brk = "";
		String pakketnr = "";
		String rekeningnummer = "";
		String srt_brk_brk = "";
		String tarlgnr = "";		
		String verznr = "";
		
		for (int i = 0; i < 10; i++) {
			afvdat_brk =i+"1";	
			dlnrnr = i+"2";		
			echtsnr = i+"3";
			gebdat_brk = i+"4";
			gesl_brk = i+"5";
			nm_dlnr_brk = i+"6";
			pakketnr = i+"7";
			rekeningnummer = i+"8";
			srt_brk_brk = i+"9";
			tarlgnr = i+"10";		
			verznr = i+"11";
			
			ValueBean con=new ValueBean();
			
			con.setDlnrnr(dlnrnr);
			con.setAfvdat_brk(afvdat_brk);
			con.setEchtsnr(echtsnr);
			con.setGebdat_brk(gebdat_brk);
			con.setGesl_brk(gesl_brk);
			con.setNm_dlnr_brk(nm_dlnr_brk);
			con.setPakketnr(pakketnr);
			con.setRekeningnummer(rekeningnummer);
			con.setSrt_brk_brk(srt_brk_brk);
			con.setTarlgnr(tarlgnr);
			con.setVerznr(verznr);
			ArrayList<FundsVO> fundsVOs = new ArrayList<FundsVO>();
			for (int j = 0; j < 10; j++) {
				FundsVO fundVO = new FundsVO();
				fundVO.setEenheden( ""+i+j+"12");
				fundVO.setWaarde(""+i+j+"13");
				fundVO.setVerkoopkosten(""+i+j+"14");
				fundsVOs.add(fundVO); 
			}
			con.setFund(fundsVOs);
			
						
			contractList.add(con);
		}

		ArrayList<ValueBean> header = new ArrayList<ValueBean>();	
		header.add(new ValueBean("88888", "18/01/1987", "01/02/1956","26/08/1954"));

		ArrayList<Columns> cols = new ArrayList<Columns>();	
		for (int j = 0; j < 10; j++) {
			cols.add(new Columns("001"+j,"18"+j+"€"));
		}
	
		Map map = new HashMap();
		map.put("records", contractList);
		map.put("header", header);
		map.put("cols", cols);
    	    
	    XLSTransformer transformer = new XLSTransformer();
	    transformer.transformXLS(templateFileName, map, destFileName);

	}

}
