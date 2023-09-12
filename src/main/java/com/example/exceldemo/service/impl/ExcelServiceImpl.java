package com.example.exceldemo.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.example.exceldemo.entity.TableOne;
import com.example.exceldemo.entity.TableTwo;
import com.example.exceldemo.service.ExcelService;
import com.example.exceldemo.utils.ImportExcelHelper;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.log4j.Log4j2;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.*;
import java.util.stream.Collectors;


@Service
@Log4j2
public class ExcelServiceImpl implements ExcelService {


    /**
     * 处理数据
     *
     * @param multipartFile 文件
     */
    @Override
    public void handelData(MultipartFile multipartFile, HttpServletResponse response, String filterStr) throws IOException {
        List<TableOne> notFoundList = new ArrayList<>();

        List<TableTwo> reList = new ArrayList<>();

        ImportExcelHelper<TableOne> helper = new ImportExcelHelper<>();
        List<TableOne> tableOneList = helper.getList(multipartFile, TableOne.class, 0, 1);


        ImportExcelHelper<TableTwo> helper1 = new ImportExcelHelper<>();
        List<TableTwo> tableTwoList = helper1.getList(multipartFile, TableTwo.class, 1, 1);


        //根据模型id分组
        Map<String, List<TableTwo>> tableTwoMap = tableTwoList.stream().filter(t->t.getGlId()!=null).collect(Collectors.groupingBy(TableTwo::getGlId));

        Map<String, List<TableTwo>> tableTwoHandMap = new HashMap<>();
        for (Map.Entry<String, List<TableTwo>> entry : tableTwoMap.entrySet()) {
            //判断一组里面是否有多个externalid
            List<String> externalList = entry.getValue().stream().map(TableTwo::getExternalId).distinct().toList();
            if (externalList.size() > 1) {
                reList.addAll(entry.getValue());
                continue;
            }



            //找到字段名称是编码的数据，当作新key
            List<TableTwo> keyList = entry.getValue().stream().filter(t -> "编码".equals(t.getPropertyName())).toList();
            if (keyList.size() > 0) {
                List<TableTwo> isPersent = tableTwoHandMap.get(keyList.get(0).getValue());

                if (CollectionUtils.isEmpty(isPersent)) {
                    tableTwoHandMap.put(keyList.get(0).getValue(), entry.getValue());
                } else {
                    isPersent.addAll(entry.getValue());
                }

            }

        }


        //如果有B表错误 直接返回
        if (!CollectionUtils.isEmpty(reList)) {
            EasyExcel.write(getExcelOutputStream(response,"B表问题数据"), TableTwo.class).sheet("B表").doWrite(reList);
            return;
        }



        //循环A表 把A表的 构建构建名称、全路径匹配 匹配不到放入返回list
        for (TableOne one : tableOneList) {
            //是否指定过滤某些构件，如果为参数null则跳过此判断
            if (filterStr != null && !one.getMeberPath().contains(filterStr)) {
                continue;
            }

            List<TableTwo> twoList = tableTwoHandMap.get(one.getMemberCode());
            if (CollectionUtils.isEmpty(twoList)){
                notFoundList.add(one);
                continue;
            }

            //查询名称和路径是否都符合
            List<TableTwo> twoName = twoList.stream()
                    .filter(t -> "分项名称".equals(t.getPropertyName()) && one.getMemberName().equals(t.getValue()))
                    .toList();
            List<TableTwo> twoPath = twoList.stream()
                    .filter(t -> "路径".equals(t.getPropertyName()) && one.getMeberPath().equals(t.getValue()))
                    .toList();
            if (CollectionUtils.isEmpty(twoName) || CollectionUtils.isEmpty(twoPath)){
                notFoundList.add(one);
                continue;
            }

            List<String> twoNameGlId = twoName.stream().map(TableTwo::getGlId).toList();
            List<String> twoPathGlId = twoPath.stream().map(TableTwo::getGlId).toList();

            List<String> retain = twoNameGlId.stream().filter(twoPathGlId::contains).toList();
            if (CollectionUtils.isEmpty(retain)) {
                notFoundList.add(one);
            }
        }
        
        //写返回excel
        String fileName = Optional.ofNullable(filterStr).orElse("") + "未匹配数据" + System.currentTimeMillis();
        EasyExcel.write(getExcelOutputStream(response,fileName), TableOne.class).sheet("未匹配").doWrite(notFoundList);

    }


    @Override
    public void handRepeatB(MultipartFile multipartFile, HttpServletResponse response) throws IOException {
        ImportExcelHelper<TableTwo> helper1 = new ImportExcelHelper<>();
        List<TableTwo> tableTwoList = helper1.getList(multipartFile, TableTwo.class, 1);

        List<TableTwo> reList;


        reList = tableTwoList.stream().filter(t->t.getGlId() == null).collect(Collectors.toList());
        //根据模型id分组
        Map<String, List<TableTwo>> tableTwoMap = tableTwoList.stream().filter(t->t.getGlId()!=null).collect(Collectors.groupingBy(TableTwo::getGlId));

        for (Map.Entry<String, List<TableTwo>> entry : tableTwoMap.entrySet()) {
            //判断一组里面是否有多个externalid
            List<String> externalList = entry.getValue().stream().map(TableTwo::getExternalId).distinct().toList();
            if (externalList.size() > 1) {
                reList.addAll(entry.getValue());
            }
        }
        //如果有B表错误 直接返回
        if (!CollectionUtils.isEmpty(reList)) {
            EasyExcel.write(getExcelOutputStream(response,"B表问题数据"), TableTwo.class).sheet("B表").doWrite(reList);
        } else {
            EasyExcel.write(getExcelOutputStream(response,"B表整合"), TableTwo.class).sheet("B表").doWrite(tableTwoList);
        }

    }

    public static OutputStream getExcelOutputStream(HttpServletResponse response, String fileName) {
        try {
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8") + ".xlsx");
            response.setHeader("Access-Control-Expose-Headers", "requestType,Content-Disposition");
            response.setHeader("requestType", "file");
            return response.getOutputStream();
        } catch (IOException e) {
            log.error(e.getMessage(), e);
        }
        return null;
    }

    @Override
    public void handMergeB(MultipartFile multipartFile, HttpServletResponse response) throws IOException {
        ImportExcelHelper<TableTwo> helper1 = new ImportExcelHelper<>();
        List<TableTwo> tableTwoList = helper1.getList(multipartFile, TableTwo.class, 1, 1);


        List<TableTwo> returnList = new ArrayList<>();
        //根据模型id分组
        Map<String, List<TableTwo>> tableTwoMap = tableTwoList.stream().filter(t->t.getGlId()!=null).collect(Collectors.groupingBy(TableTwo::getGlId));

        for (Map.Entry<String, List<TableTwo>> entry : tableTwoMap.entrySet()) {
            //判断一组里面是否有多个编码
            List<TableTwo> codeList = entry.getValue().stream().filter(t->"编码".equals(t.getPropertyName())).collect(Collectors.toList());

            if (codeList.size() > 1) {
                returnList.addAll(entry.getValue());
            }

        }

        //如果有B表错误 直接返回
        if (!CollectionUtils.isEmpty(returnList)) {
            EasyExcel.write(getExcelOutputStream(response,"B表问题数据"), TableTwo.class).sheet("B表").doWrite(returnList);
        }
    }
}
