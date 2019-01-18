
//所使用的jar包--poi （ jxl对表格类型有限制）

//这是针对我的本地Excel表格是有规律的，这样操作，读取的数据好封装，然后方便去在数据库操作。
/**
 * 从本地excel表格中解析数据
 */
public static List<City> readExcel(String path) {
        List<City> list = null;  //这是自己将数据封装的对象
        try {
            File excelFile = new File(path); //创建文件对象
            FileInputStream is = new FileInputStream(excelFile); //文件流
            Workbook workbook = WorkbookFactory.create(is);    //这种方式 Excel 2003/2007/2010
            int sheetCount = workbook.getNumberOfSheets();
            list = new ArrayList<City>(); //存储数据容器
            for (int s = 0; s < sheetCount; s++) {
                Sheet sheet = workbook.getSheetAt(s);
                int rowCount = sheet.getPhysicalNumberOfRows(); //获取总行数   
                System.out.println("行:"+rowCount);
                for (int r = 1; r < rowCount; r++) {
                    Row row = sheet.getRow(r);
                    row.getCell(0).setCellType(CellType.STRING);  //因为表格中存在数据，所以设置成string类型，再去读取
                    row.getCell(1).setCellType(CellType.STRING);
                    String areCode = row.getCell(0).getStringCellValue();
                    String id1 = row.getCell(1).getStringCellValue();
                    String id = id1.trim();
                    City city = new City();
                    city.setId(Integer.valueOf(id));
                    city.setAreaCode(areCode);
                    System.out.println("id" + city.getId() +";areCode"+city.getAreaCode());
                    list.add(city);
                }
            }
            is.close();
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }
