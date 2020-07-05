package com.antonioalejandro.utils.excel;

public class ExcelBooleanData {

	private boolean booleanValue;
	private String textTrueOption;
	private String textFalseOption;

	public boolean getBooleanValue() {
		return booleanValue;
	}

	public void setBooleanValue(final boolean booleanValue) {
		this.booleanValue = booleanValue;
	}

	public String getTextTrueOption() {
		return textTrueOption;
	}

	public void setTextTrueOption(final String textTrueOption) {
		this.textTrueOption = textTrueOption;
	}

	public String getTextFalseOption() {
		return textFalseOption;
	}

	public void setTextFalseOption(final String textFalseOption) {
		this.textFalseOption = textFalseOption;
	}

	public ExcelBooleanData(final boolean booleanValue, final String textTrueOption, final String textFalseOption) {
		super();
		this.booleanValue = booleanValue;
		this.textTrueOption = textTrueOption;
		this.textFalseOption = textFalseOption;
	}

	public ExcelBooleanData(final boolean value) {
		this.booleanValue = value;
		this.textFalseOption = "False";
		this.textTrueOption = "True";
	}

}
