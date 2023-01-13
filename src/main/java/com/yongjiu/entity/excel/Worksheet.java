package com.yongjiu.entity.excel;

import lombok.Data;

import java.util.List;

@Data
public class Worksheet {

	private String Name;

	private Table table;
	private List<DataValidations> dataValidationList;

}
