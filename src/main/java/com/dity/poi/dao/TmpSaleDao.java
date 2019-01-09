package com.dity.poi.dao;

import com.dity.poi.pojo.TmpSale;

import java.util.List;
import java.util.Map;

/**
 * @author:yuhang
 * @Date:2018/5/2
 */
public interface TmpSaleDao {

    void insertBatch(List<TmpSale> list);

    List<Map<String,Object>> queryResult();

    List<TmpSale> queryAll();
}
