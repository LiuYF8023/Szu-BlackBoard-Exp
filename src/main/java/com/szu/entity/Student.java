package com.szu.entity;

import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;

@Data
public class Student {
	@ColumnWidth(30)
	private String name;
	private String exp1;
	private String exp2;
	private String exp3;
	private String exp4;
	private String exp51;
	private String exp52;
	private String exp53;
	private String exp6;
	private String exp7;
	private String exp8;
	private String exp9;
}
