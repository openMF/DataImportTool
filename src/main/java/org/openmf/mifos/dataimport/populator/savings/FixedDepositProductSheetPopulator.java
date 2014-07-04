package org.openmf.mifos.dataimport.populator.savings;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Currency;
import org.openmf.mifos.dataimport.dto.savings.FixedDepositProduct;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.openmf.mifos.dataimport.populator.AbstractWorkbookPopulator;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

public class FixedDepositProductSheetPopulator extends AbstractWorkbookPopulator {
	
	private static final Logger logger = LoggerFactory.getLogger(FixedDepositProductSheetPopulator.class);
	
    private final RestClient client;
	
	private String content;
	
	private static final int ID_COL = 0;
	private static final int NAME_COL = 1;
	private static final int SHORT_NAME_COL = 2;
	private static final int NOMINAL_ANNUAL_INTEREST_RATE_COL = 3;
	private static final int INTEREST_COMPOUNDING_PERIOD_COL = 4;
	private static final int INTEREST_POSTING_PERIOD_COL = 5;
	private static final int INTEREST_CALCULATION_COL = 6;
	private static final int INTEREST_CALCULATION_DAYS_IN_YEAR_COL = 7;
	private static final int LOCKIN_PERIOD_COL = 8;
	private static final int LOCKIN_PERIOD_FREQUENCY_COL = 9;
	private static final int CURRENCY_COL = 10;
	private static final int MIN_DEPOSIT_COL = 11;
	private static final int MAX_DEPOSIT_COL = 12;
	private static final int DEPOSIT_COL = 13;
	private static final int MIN_DEPOSIT_TERM_COL = 14;
	private static final int MIN_DEPOSIT_TERM_TYPE_COL = 15;
	private static final int MAX_DEPOSIT_TERM_COL = 16;
	private static final int MAX_DEPOSIT_TERM_TYPE_COL = 17;
	private static final int PRECLOSURE_PENAL_APPLICABLE_COL = 18;
	private static final int PRECLOSURE_PENAL_INTEREST_COL = 19;
	private static final int PRECLOSURE_INTEREST_TYPE_COL = 20;
	private static final int IN_MULTIPLES_OF_DEPOSIT_TERM_COL = 21;
	private static final int IN_MULTIPLES_OF_DEPOSIT_TERM_TYPE_COL = 22;
	
	private List<FixedDepositProduct> products;
	
	public FixedDepositProductSheetPopulator(RestClient client) {
        this.client = client;
    }
	
	@Override
    public Result downloadAndParse() {
    	Result result = new Result();
        try {
        	client.createAuthToken();
        	products = new ArrayList<FixedDepositProduct>();
            content = client.get("fixeddepositproducts");
            Gson gson = new Gson();
            JsonElement json = new JsonParser().parse(content);
            JsonArray array = json.getAsJsonArray();
            Iterator<JsonElement> iterator = array.iterator();
            while(iterator.hasNext()) {
            	json = iterator.next();
            	FixedDepositProduct product = gson.fromJson(json, FixedDepositProduct.class);
            	products.add(product);
            }
        } catch (Exception e) {
            result.addError(e.getMessage());
            logger.error(e.getMessage());
        }
        return result;
    }
	
	@Override
	 public Result populate(Workbook workbook) {
	    	Result result = new Result();
	    	try{
	    		int rowIndex = 1;
	            Sheet productSheet = workbook.createSheet("Products");
	            setLayout(productSheet);
	            CellStyle dateCellStyle = workbook.createCellStyle();
	            short df = workbook.createDataFormat().getFormat("dd-mmm");
	            dateCellStyle.setDataFormat(df);
	            for(FixedDepositProduct product : products) {
	            	Row row = productSheet.createRow(rowIndex++);
	            	writeInt(ID_COL, row, product.getId());
	            	writeString(NAME_COL, row, product.getName().trim().replaceAll("[ )(]", "_"));
	            	writeString(SHORT_NAME_COL, row, product.getShortName().trim().replaceAll("[ )(]", "_"));
	            	writeDouble(NOMINAL_ANNUAL_INTEREST_RATE_COL, row, product.getNominalAnnualInterestRate());
	            	writeString(INTEREST_COMPOUNDING_PERIOD_COL, row, product.getInterestCompoundingPeriodType().getValue());
	            	writeString(INTEREST_POSTING_PERIOD_COL, row, product.getInterestPostingPeriodType().getValue());
	            	writeString(INTEREST_CALCULATION_COL, row, product.getInterestCalculationType().getValue());
	            	writeString(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, row, product.getInterestCalculationDaysInYearType().getValue());
	            	writeDouble(DEPOSIT_COL, row, product.getDepositAmount());
	            	writeString(PRECLOSURE_PENAL_APPLICABLE_COL, row, product.getPreClosurePenalApplicable());
	            	writeInt(MIN_DEPOSIT_TERM_COL, row, product.getMinDepositTerm());
	            	writeString(MIN_DEPOSIT_TERM_TYPE_COL, row, product.getMinDepositTermType().getValue());
	            	
	            	if(product.getMinDepositAmount() != null)
	            		writeDouble(MIN_DEPOSIT_COL, row, product.getMinDepositAmount());
	            	if(product.getMaxDepositAmount() != null)
	            		writeDouble(MAX_DEPOSIT_COL, row, product.getMaxDepositAmount());
	            	if(product.getMaxDepositTerm() != null)
	            		writeInt(MAX_DEPOSIT_TERM_COL, row, product.getMaxDepositTerm());
	            	if(product.getInMultiplesOfDepositTerm() != null)
	            		writeInt(IN_MULTIPLES_OF_DEPOSIT_TERM_COL, row, product.getInMultiplesOfDepositTerm());
	            	if(product.getPreClosurePenalInterest() != null)
	            		writeDouble(PRECLOSURE_PENAL_INTEREST_COL, row, product.getPreClosurePenalInterest());
	            	if(product.getMaxDepositTermType() != null)
	            		writeString(MAX_DEPOSIT_TERM_TYPE_COL, row, product.getMaxDepositTermType().getValue());
	            	if(product.getPreClosureInterestOnType() != null)
	            		writeString(PRECLOSURE_INTEREST_TYPE_COL, row, product.getPreClosureInterestOnType().getValue());
	            	if(product.getInMultiplesOfDepositTermType() != null)
	            		writeString(IN_MULTIPLES_OF_DEPOSIT_TERM_TYPE_COL, row, product.getInMultiplesOfDepositTermType().getValue());
	            	
	            	if(product.getLockinPeriodFrequency() != null)
	            	    writeInt(LOCKIN_PERIOD_COL, row, product.getLockinPeriodFrequency());
	            	if(product.getLockinPeriodFrequencyType() != null)
	            	    writeString(LOCKIN_PERIOD_FREQUENCY_COL, row, product.getLockinPeriodFrequencyType().getValue());
	            	Currency currency = product.getCurrency();
	            	writeString(CURRENCY_COL, row, currency.getCode());
	            }
//	        	productSheet.protectSheet("");
   	} catch (RuntimeException re) {
   		result.addError(re.getMessage());
   		logger.error(re.getMessage());
   	}
       return result;
     }
	
	private void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        worksheet.setColumnWidth(ID_COL, 2000);
        worksheet.setColumnWidth(NAME_COL, 5000);
        worksheet.setColumnWidth(SHORT_NAME_COL, 2000);
        worksheet.setColumnWidth(NOMINAL_ANNUAL_INTEREST_RATE_COL, 2000);
        worksheet.setColumnWidth(INTEREST_COMPOUNDING_PERIOD_COL, 3000);
        worksheet.setColumnWidth(INTEREST_POSTING_PERIOD_COL, 3000);
        worksheet.setColumnWidth(INTEREST_CALCULATION_COL, 3000);
        worksheet.setColumnWidth(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, 3000);
        worksheet.setColumnWidth(LOCKIN_PERIOD_COL, 3000);
        worksheet.setColumnWidth(LOCKIN_PERIOD_FREQUENCY_COL, 3000);
        worksheet.setColumnWidth(CURRENCY_COL, 3000);
        worksheet.setColumnWidth(MIN_DEPOSIT_COL, 3000);
        worksheet.setColumnWidth(MAX_DEPOSIT_COL, 3000);
        worksheet.setColumnWidth(DEPOSIT_COL, 3000);
        worksheet.setColumnWidth(MIN_DEPOSIT_TERM_COL, 3000);
        worksheet.setColumnWidth(MAX_DEPOSIT_TERM_COL, 3000);
        worksheet.setColumnWidth(MIN_DEPOSIT_TERM_TYPE_COL, 3000);
        worksheet.setColumnWidth(MAX_DEPOSIT_TERM_TYPE_COL, 3000);
        worksheet.setColumnWidth(PRECLOSURE_PENAL_APPLICABLE_COL, 3000);
        worksheet.setColumnWidth(PRECLOSURE_PENAL_INTEREST_COL, 3000);
        worksheet.setColumnWidth(PRECLOSURE_INTEREST_TYPE_COL, 3000);
        worksheet.setColumnWidth(IN_MULTIPLES_OF_DEPOSIT_TERM_COL, 3000);
        worksheet.setColumnWidth(IN_MULTIPLES_OF_DEPOSIT_TERM_TYPE_COL, 3000);
        
        writeString(ID_COL, rowHeader, "ID");
        writeString(NAME_COL, rowHeader, "Name");
        writeString(SHORT_NAME_COL, rowHeader, "Short Name");
        writeString(NOMINAL_ANNUAL_INTEREST_RATE_COL, rowHeader, "Interest");
        writeString(INTEREST_COMPOUNDING_PERIOD_COL, rowHeader, "Interest Compounding Period");
        writeString(INTEREST_POSTING_PERIOD_COL, rowHeader, "Interest Posting Period");
        writeString(INTEREST_CALCULATION_COL, rowHeader, "Interest Calculated Using");
        writeString(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, rowHeader, "# Days In Year");
        writeString(LOCKIN_PERIOD_COL, rowHeader, "Locked In For");
        writeString(LOCKIN_PERIOD_FREQUENCY_COL, rowHeader, "Frequency");
        writeString(CURRENCY_COL, rowHeader, "Currency");
        writeString(MIN_DEPOSIT_COL, rowHeader, "Min Deposit");
        writeString(MAX_DEPOSIT_COL, rowHeader, "Max Deposit");
        writeString(DEPOSIT_COL, rowHeader, "Deposit");
        writeString(MIN_DEPOSIT_TERM_COL, rowHeader, "Min Deposit Term");
        writeString(MAX_DEPOSIT_TERM_COL, rowHeader, "Max Deposit Term");
        writeString(MIN_DEPOSIT_TERM_TYPE_COL, rowHeader, "Min Deposit Term Type");
        writeString(MAX_DEPOSIT_TERM_TYPE_COL, rowHeader, "Max Deposit Term Type");
        writeString(PRECLOSURE_PENAL_APPLICABLE_COL, rowHeader, "Preclosure Penal Applicable");
        writeString(PRECLOSURE_PENAL_INTEREST_COL, rowHeader, "Penal Interest");
        writeString(PRECLOSURE_INTEREST_TYPE_COL, rowHeader, "Penal Interest Type");
        writeString(IN_MULTIPLES_OF_DEPOSIT_TERM_COL, rowHeader, "Multiples of Deposit Term");
        writeString(IN_MULTIPLES_OF_DEPOSIT_TERM_TYPE_COL, rowHeader, "Multiples of Deposit Term Type");
	}
	
	
	public List<FixedDepositProduct> getProducts() {
		 return products;
	 }
	
	public Integer getProductsSize() {
		 return products.size();
	 }

}
