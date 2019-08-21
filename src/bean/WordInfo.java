package bean;

import java.util.Comparator;

/**
 * 提取Word内容<br></>存为WordInfo<br></> 导入Excel
 */
public class WordInfo {
    /*public WordInfo(int id, String orderNumber, String supplyOrderNumber, String powerSupply, String projectName, String goodsName, String modelName, String goodsCloums, String goodsUnit, String notwithPrice, String notwithTotalPrice, String totalPrice, String deliveryTime) {
        this.id = id;
        this.orderNumber = orderNumber;
        this.supplyOrderNumber = supplyOrderNumber;
        this.powerSupply = powerSupply;
        this.projectName = projectName;
        this.goodsName = goodsName;
        this.modelName = modelName;
        this.goodsCloums = goodsCloums;
        this.goodsUnit = goodsUnit;
        this.notwithPrice = notwithPrice;
        this.notwithTotalPrice = notwithTotalPrice;
        this.totalPrice = totalPrice;
        this.deliveryTime = deliveryTime;
    }*/
    public static class OrderComparator implements Comparator<WordInfo> {

        @Override
        public int compare(WordInfo o1, WordInfo o2) {
            int i1 = Integer.valueOf(o1.getOrderNumber());
            int i2 = Integer.valueOf(o2.getOrderNumber());
            if (i1 < i2) {
                return 1;
            } else if (i1 > i2) {
                return -1;
            } else {
                return 0;
            }
        }

    }

    @Override
    public String toString() {
        return "源文件" + srcfile +
                "序号" + id +
                " | 订单编号" + orderNumber +
                " | 供货单号" + supplyOrderNumber +
                " | 供电局" + powerSupply +
                " | 项目名称" + projectName +
                " | 货物名称" + goodsName +
                " | 型号" + modelName +
                " | 数量" + goodsCloums +
                " | 单位" + goodsUnit +
                " | 含税单价" + notwithPrice +
                " | 含税总价" + notwithTotalPrice +
                " | 总价" + totalPrice +
                " | 交货时间" + deliveryTime;
    }

    private int id;//序号
    private String srcfile;//源文件
    private String orderNumber;//订单编号
    private String supplyOrderNumber;//供货单号
    private String powerSupply;//供电局
    private String projectName;//项目名称
    private String goodsName;//货物名称
    private String modelName;//型号
    private String goodsCloums;//数量
    private String goodsUnit;//单位
    private String notwithPrice;//不含税单价
    private String notwithTotalPrice;//不含税总价
    private String totalPrice;//含税总价
    private String deliveryTime;//交货时间

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getSrcfile() {
        return srcfile;
    }

    public void setSrcfile(String srcfile) {
        this.srcfile = srcfile;
    }

    public String getOrderNumber() {
        return orderNumber;
    }

    public void setOrderNumber(String orderNumber) {
        this.orderNumber = orderNumber;
    }

    public String getSupplyOrderNumber() {
        return supplyOrderNumber;
    }

    public void setSupplyOrderNumber(String supplyOrderNumber) {
        this.supplyOrderNumber = supplyOrderNumber;
    }

    public String getPowerSupply() {
        return powerSupply;
    }

    public void setPowerSupply(String powerSupply) {
        this.powerSupply = powerSupply;
    }

    public String getProjectName() {
        return projectName;
    }

    public void setProjectName(String projectName) {
        this.projectName = projectName;
    }

    public String getGoodsName() {
        return goodsName;
    }

    public void setGoodsName(String goodsName) {
        this.goodsName = goodsName;
    }

    public String getModelName() {
        return modelName;
    }

    public void setModelName(String modelName) {
        this.modelName = modelName;
    }

    public String getGoodsCloums() {
        return goodsCloums;
    }

    public void setGoodsCloums(String goodsCloums) {
        this.goodsCloums = goodsCloums;
    }

    public String getGoodsUnit() {
        return goodsUnit;
    }

    public void setGoodsUnit(String goodsUnit) {
        this.goodsUnit = goodsUnit;
    }

    public String getNotwithPrice() {
        return notwithPrice;
    }

    public void setNotwithPrice(String notwithPrice) {
        this.notwithPrice = notwithPrice;
    }

    public String getNotwithTotalPrice() {
        return notwithTotalPrice;
    }

    public void setNotwithTotalPrice(String notwithTotalPrice) {
        this.notwithTotalPrice = notwithTotalPrice;
    }

    public String getTotalPrice() {
        return totalPrice;
    }

    public void setTotalPrice(String totalPrice) {
        this.totalPrice = totalPrice;
    }

    public String getDeliveryTime() {
        return deliveryTime;
    }

    public void setDeliveryTime(String deliveryTime) {
        this.deliveryTime = deliveryTime;
    }
}
