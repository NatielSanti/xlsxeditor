package ru.natiel.xlsxeditor.dto;

public class ImsiData {
    private String brandCd;
    private String requestChannelCd;
    private String costCenter;

    public ImsiData(String brandCd, String requestChannelCd, String costCenter){
        this.brandCd = brandCd;
        this.requestChannelCd = requestChannelCd;
        this.costCenter = costCenter;
    }

    public String getBrandCd() {
        return brandCd;
    }

    public void setBrandCd(String brandCd) {
        this.brandCd = brandCd;
    }

    public String getRequestChannelCd() {
        return requestChannelCd;
    }

    public void setRequestChannelCd(String requestChannelCd) {
        this.requestChannelCd = requestChannelCd;
    }

    public String getCostCenter() {
        return costCenter;
    }

    public void setCostCenter(String costCenter) {
        this.costCenter = costCenter;
    }
}
