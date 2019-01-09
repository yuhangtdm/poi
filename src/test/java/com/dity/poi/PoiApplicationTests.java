package com.dity.poi;

import com.dity.poi.dao.TmpSaleDao;
import com.dity.poi.pojo.TmpSale;
import com.dity.poi.util.ExcelUtil;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

	@Autowired
	TmpSaleDao tmpSaleDao;

	@Test
	public void contextLoads() {
	}

	@Test
	public void testRead(){
		ExcelUtil.initByPath("D:\\city.xls");
		ExcelUtil.readExcel();
		System.out.println(ExcelUtil.list.size());
		tmpSaleDao.insertBatch(ExcelUtil.list);
	}

	@Test
	public void testUpdate(){
		List<TmpSale> tmpSales = tmpSaleDao.queryAll();
		File file=new File("D:\\tmp.sql");
		BufferedWriter bw=null;
		try {
			bw=new BufferedWriter(new FileWriter(file));
			for(TmpSale tmpSale:tmpSales){
				bw.write("update sa_sale_info set employee_no='"+tmpSale.getEmployeeNo()+"' where idno='"+tmpSale.getIdno()+"';" );
				bw.newLine();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}finally {
			if(bw!=null){
				try {
					bw.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}



	}

}
