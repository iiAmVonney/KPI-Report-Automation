import java.util.Date;

public class tDate {
	
	
	
	public tDate(String g, long d, String des)
	{
		gid = g;
		date = new Date(d);
		desc = des;
	}
	
	public String getGid()
	{
		return gid;
	}
	
	public Date getDate()
	{
		return date;
	}
	
	public void setDate(long d)
	{
		date = new Date(d);
	}
	
	public String toString()
	{
		if( this==null)
			return "null";
		else
			return desc;
	}
	
	String gid,desc;
	Date date;

}
