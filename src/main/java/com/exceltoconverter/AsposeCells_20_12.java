package com.exceltoconverter;

import javassist.ClassPool;
import javassist.CtClass;

public class AsposeCells_20_12 {
    public static void main(String[] args) throws Exception {
        ClassPool pool = ClassPool.getDefault();
        pool.insertClassPath("aspose-cells-20.12.jar");
        
        // 处理Workbook类
        CtClass clazz1 = pool.get("com.aspose.cells.Workbook");
        clazz1.getClassInitializer().setBody("{" +
                "    com.aspose.cells.License license = new com.aspose.cells.License();" +
                "    license.setLicense(new java.io.StringReader(\"<License> <Data> <Products> <Product>Aspose.Cells for Java</Product> </Products> <EditionType>Enterprise</EditionType> <SubscriptionExpiry>29991231</SubscriptionExpiry> <LicenseExpiry>29991231</LicenseExpiry> <SerialNumber>evilrule</SerialNumber> </Data> <Signature>evilrule</Signature> </License>\"));" +
                "}");
        clazz1.writeFile();
        
        // 处理License类
        CtClass clazz = pool.getCtClass("com.aspose.cells.License");
        clazz.getDeclaredMethod("isLicenseSet").setBody("{return true;}");
        clazz.getDeclaredMethod("a", new CtClass[]{pool.get("java.lang.String"), pool.get("java.lang.String"), pool.get("boolean"), pool.get("boolean")}).setBody("{return true;}");
        clazz.getDeclaredMethod("l", new CtClass[]{pool.get("java.lang.String")}).setBody("{return new java.util.Date(Long.MAX_VALUE);}");
        clazz.getDeclaredMethod("k", new CtClass[]{pool.get("java.lang.String")}).setBody("{return true;}");
        clazz.writeFile();
        
        System.out.println("Aspose Cells 20.12 license patched successfully!");
    }
}