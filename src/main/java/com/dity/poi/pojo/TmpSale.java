package com.dity.poi.pojo;

import lombok.Data;

/**
 * @author:yuhang
 * @Date:2018/5/2
 */
@Data
public class TmpSale{
    private String employeeNo;
    private String idno;

    public TmpSale(){

    }

    public TmpSale(String employeeNo, String idno) {
        this.employeeNo = employeeNo;
        this.idno = idno;
    }

}