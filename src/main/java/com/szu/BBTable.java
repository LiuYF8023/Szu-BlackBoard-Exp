package com.szu;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.szu.entity.Student;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.util.*;

/**
 * Hello world!
 */
public class BBTable {
	private static final String url = "https://elearning.szu.edu.cn/webapps/gradebook/do/instructor/viewNeedsGrading?sortCol=attemptDate&sortDir=ASCENDING&showAll=true&editPaging=false&course_id=_19018_1&startIndex=0";

	private static final String cookieSzu = "cookie: JSESSIONID=7F2119F684A55584EEE21A5CB820B128; COOKIE_CONSENT_ACCEPTED=true; s_session_id=F2E430846EB982AE19900668D05F8E46; JSESSIONID=4F7C4CA9C96FAE9C273FAA693BC7DCEF; CdnSignedValidation=false; BbClientCalenderTimeZone=Asia/Shanghai; BbClientDownloadExecuting=false; web_client_cache_guid=6fbcf829-a3cf-4ffb-94fa-dd8f1423f9b7; xythosdrive=0";
	private static final String fileName = "src/main/resources/Experiment_Statistics.xlsx";

	public static void main(String[] args) throws Exception {
		BBTable.run();
	}

	public static void run() throws Exception {
		StringBuffer data = new StringBuffer();
		try {
			// 根据URL获取当前URL界面的doc对象，里面存储着界面的所有元素，类似于BOM
			// 手动设置cookies
			Document doc = Jsoup.connect(url).header("Cookie", cookieSzu).get();

			// 表格大小
			int tableSize = Integer.parseInt(doc.getElementsByClass("criteriaSummary").text().split(" ")[1]);
			// 每一个学生都对应一个对象
			List<Student> list = ListUtils.newArrayList();

			// 标记哪些学生已经被获取了
			Set<String> itered = new HashSet<>();
			// 循环获取数据

			for (int i = 0; i < tableSize; i++) {

				// 获取是第几次提交
				String submitTimeTemp = doc.getElementById("listContainer_row:" + i).select(".table-data-cell-value").text().split(" ")[1].split("_")[0];
				String submitTime = submitTimeTemp.substring(2, submitTimeTemp.length());

				// 是否逾期
				Element overTime = doc.getElementById("listContainer_row:" + i);
				boolean overTimeBoolean = overTime.select(".table-data-cell-value").select(".lateIndicator").text().equals("逾期");

				// 获取名字
				String name = doc.getElementById("listContainer_row:" + i).select(".gradeAttempt").text().split(" ")[1];

				Student stemp = null;
				// 如果这个已经包含在内了，那么我直接跳过就行了，说明前面已经找过了
				if (itered.contains(name)) {
					continue;
				} else {
					// 创建该学生对应的对象 然后修改实验提交情况
					stemp = new Student();
					stemp.setName(name);
					itered.add(name);

					if (name.equals(name)) {
						if (submitTime.equals("1")) {
							stemp.setExp1("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("2")) {
							stemp.setExp2("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("3")) {
							stemp.setExp3("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("4")) {
							stemp.setExp4("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("5_1")) {
							stemp.setExp51("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("5_2")) {
							stemp.setExp52("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("5_3")) {
							stemp.setExp53("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("6")) {
							stemp.setExp6("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("7")) {
							stemp.setExp7("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("8")) {
							stemp.setExp8("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("9")) {
							stemp.setExp9("√" + (overTimeBoolean ? " 逾期" : ""));
						}
					}
				}

				for (int j = i + 1; j < tableSize; j++) {
					// 获取是第几次提交
					submitTimeTemp = doc.getElementById("listContainer_row:" + j).select(".table-data-cell-value").text().split(" ")[1].split("_")[0];
					submitTime = submitTimeTemp.substring(2, submitTimeTemp.length());

					// 是否逾期
					overTime = doc.getElementById("listContainer_row:" + i);
					overTimeBoolean = overTime.select(".table-data-cell-value").select(".lateIndicator").text().equals("逾期");

					// 获取后面的相同名字
					String nameAfter = doc.getElementById("listContainer_row:" + j).select(".gradeAttempt").text().split(" ")[1];
					if (name.equals(nameAfter)) {
						if (submitTime.equals("1")) {
							stemp.setExp1("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("2")) {
							stemp.setExp2("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("3")) {
							stemp.setExp3("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("4")) {
							stemp.setExp4("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("5_1")) {
							stemp.setExp51("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("5_2")) {
							stemp.setExp52("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("5_3")) {
							stemp.setExp53("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("6")) {
							stemp.setExp6("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("7")) {
							stemp.setExp7("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("8")) {
							stemp.setExp8("√" + (overTimeBoolean ? " 逾期" : ""));
						} else if (submitTime.equals("9")) {
							stemp.setExp9("√" + (overTimeBoolean ? " 逾期" : ""));
						}
					}
				}
				list.add(stemp);
			}
			try (ExcelWriter excelWriter = EasyExcel.write(fileName, Student.class).build()) {
				WriteSheet writeSheet = EasyExcel.writerSheet("expResult").build();
				excelWriter.write(list, writeSheet);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			System.out.println("导出完成");
		}
	}

}
