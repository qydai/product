package exportXLS.entity;

import java.text.SimpleDateFormat;
import java.util.Date;

import exportXLS.annotation.Column;
import exportXLS.annotation.Xls;

@Xls(name="个人身份信息")
public class UserInfo {
	@Column(name = "序号",sortNo = 1)
	private Integer id;
	@Column(name = "姓名",sortNo = 2)
	private String name;
	@Column(name = "生日",sortNo = 7, formatter = "dateFormat")
	private Date brithDate;
	@Column(name = "电话",sortNo = 4)
	private String phone;
	@Column(name = "地址",sortNo = 5)
	private String address;
	@Column(name = "QQ",sortNo = 6)
	private String qq;
	@Column(name = "性别",sortNo = 3, formatter = "sexFormat")
	private Integer sex;
	@Column(name = "身份证",sortNo = 8)
	private String idCard;
	
	
	public void setId(Integer id) {
		this.id = id;
	}
	
	public void setName(String name) {
		this.name = name;
	}
	
	public void setBrithDate(Date brithDate) {
		this.brithDate = brithDate;
	}
	
	public void setPhone(String phone) {
		this.phone = phone;
	}
	
	public void setAddress(String address) {
		this.address = address;
	}
	
	public void setQq(String qq) {
		this.qq = qq;
	}
	
	public void setSex(Integer sex) {
		this.sex = sex;
	}
	
	public void setIdCard(String idCard) {
		this.idCard = idCard;
	}
	
	public Integer getId() {
		return id;
	}
	public String getName() {
		return name;
	}
	public Date getBrithDate() {
		return brithDate;
	}
	public String getPhone() {
		return phone;
	}
	public String getAddress() {
		return address;
	}
	public String getQq() {
		return qq;
	}
	public Integer getSex() {
		return sex;
	}
	public String getIdCard() {
		return idCard;
	}
	public String dateFormat(Date date){
		return new SimpleDateFormat("yyyy-MM-dd").format(date);
	}
	public String sexFormat(Integer sex){
		return sex == 0 ? "男" : "女"; 
	}
}
