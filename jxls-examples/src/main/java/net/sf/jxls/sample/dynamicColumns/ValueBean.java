package net.sf.jxls.sample.dynamicColumns;


import java.util.Collection;

public class ValueBean {
    
	public ValueBean() {
		
	}	
		
	public ValueBean(String Contractnummer,String Valutadatumvan,String Valutadatumtot,String Draaidatum ) {
        this.Contractnummer = Contractnummer;
        this.Valutadatumvan = Valutadatumvan;
        this.Valutadatumtot = Valutadatumtot;
        this.Draaidatum = Draaidatum;
    }
	//header data start
	private String Contractnummer;
	private String Valutadatumvan;
	private String Valutadatumtot;
	private String Draaidatum;	
	//header data finish
	
	//records data start	
	private String rekeningnummer;
	private String dlnrnr;
	private String pakketnr;
	private String echtsnr;
	private String verznr;
	private String tarlgnr;
	private String srt_brk_brk;
	private String nm_dlnr_brk;
	private String gesl_brk;
	private String gebdat_brk;
	private String afvdat_brk;	
	private Collection<FundsVO> fund;
	//records data finish
	
	public String getAfvdat_brk() {
		return afvdat_brk;
	}

	public void setAfvdat_brk(String afvdat_brk) {
		this.afvdat_brk = afvdat_brk;
	}

	public String getContractnummer() {
		return Contractnummer;
	}

	public void setContractnummer(String contractnummer) {
		Contractnummer = contractnummer;
	}

	public String getDlnrnr() {
		return dlnrnr;
	}

	public void setDlnrnr(String dlnrnr) {
		this.dlnrnr = dlnrnr;
	}

	public String getDraaidatum() {
		return Draaidatum;
	}

	public void setDraaidatum(String draaidatum) {
		Draaidatum = draaidatum;
	}

	public String getEchtsnr() {
		return echtsnr;
	}

	public void setEchtsnr(String echtsnr) {
		this.echtsnr = echtsnr;
	}

	public String getGebdat_brk() {
		return gebdat_brk;
	}

	public void setGebdat_brk(String gebdat_brk) {
		this.gebdat_brk = gebdat_brk;
	}

	public String getGesl_brk() {
		return gesl_brk;
	}

	public void setGesl_brk(String gesl_brk) {
		this.gesl_brk = gesl_brk;
	}

	public String getNm_dlnr_brk() {
		return nm_dlnr_brk;
	}

	public void setNm_dlnr_brk(String nm_dlnr_brk) {
		this.nm_dlnr_brk = nm_dlnr_brk;
	}

	public String getPakketnr() {
		return pakketnr;
	}

	public void setPakketnr(String pakketnr) {
		this.pakketnr = pakketnr;
	}

	public String getRekeningnummer() {
		return rekeningnummer;
	}

	public void setRekeningnummer(String rekeningnummer) {
		this.rekeningnummer = rekeningnummer;
	}

	public String getSrt_brk_brk() {
		return srt_brk_brk;
	}

	public void setSrt_brk_brk(String srt_brk_brk) {
		this.srt_brk_brk = srt_brk_brk;
	}

	public String getTarlgnr() {
		return tarlgnr;
	}

	public void setTarlgnr(String tarlgnr) {
		this.tarlgnr = tarlgnr;
	}

	public String getValutadatumtot() {
		return Valutadatumtot;
	}

	public void setValutadatumtot(String valutadatumtot) {
		Valutadatumtot = valutadatumtot;
	}

	public String getValutadatumvan() {
		return Valutadatumvan;
	}

	public void setValutadatumvan(String valutadatumvan) {
		Valutadatumvan = valutadatumvan;
	}

	public String getVerznr() {
		return verznr;
	}

	public void setVerznr(String verznr) {
		this.verznr = verznr;
	}
	
	public Collection<FundsVO> getFund() {
	
		return fund;
	}

	
	public void setFund(Collection<FundsVO> fund) {
	
		this.fund = fund;
	}

}
