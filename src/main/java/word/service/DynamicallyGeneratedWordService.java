package word.service;


import java.io.File;  
import java.io.FileOutputStream;  
import java.io.IOException;  
import java.io.OutputStreamWriter;  
import java.io.Writer;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;  
import java.util.HashMap;  
import java.util.List;  
import java.util.Locale;  
import java.util.Map;

import org.junit.Test;

import freemarker.cache.URLTemplateLoader;  
import freemarker.core.ParseException;  
import freemarker.template.Configuration;  
import freemarker.template.MalformedTemplateNameException;  
import freemarker.template.Template;  
import freemarker.template.TemplateException;  
import freemarker.template.TemplateNotFoundException;  
  
public class DynamicallyGeneratedWordService {  
    private static Configuration freemarkerConfig;  
      
    static {  
        freemarkerConfig = new Configuration(Configuration.VERSION_2_3_25);  
        freemarkerConfig.setEncoding(Locale.getDefault(), "UTF-8");  
    }  
      
//注意Word模板需要另存为其它格式,Word XML 2003文档才有效
// 模板使用freemarker进行遍历与获取数据操作
      
    /** 
     * 生成word文档 
     * @param filePath 
     * @throws TemplateNotFoundException 
     * @throws MalformedTemplateNameException 
     * @throws ParseException 
     * @throws IOException 
     * @throws TemplateException 
     */  
    public void genWordFile(String filePath) throws TemplateNotFoundException, MalformedTemplateNameException, ParseException, IOException, TemplateException{  
    	List stuList=new ArrayList<>();
    	for(int i=0;i<10;i++) {
    		Map student=new HashMap<>();
    		student.put("name", "张"+i);
    		student.put("age", i+10+"岁");
    		student.put("sex", "男");
    		stuList.add(student);
    	}
        Map<String,Object> result = new HashMap<String,Object>();  
        result.put("title", "是标题");  
        result.put("date", new Date());
        result.put("stuList", stuList);  
          
        freemarkerConfig.setTemplateLoader(new URLTemplateLoader() {  
              
            @Override  
            protected URL getURL(String arg0) {  
                try {
                	File file=new File("D:\\Eworkspace\\word\\src\\main\\java\\test.xml");
                	return file.toURI().toURL();
				} catch (MalformedURLException e) {
					// TODO 自动生成的 catch 块
					e.printStackTrace();
				}//此处需要注意test.xml模板的路径,不要搞错了，否则获取不到模板，我是放在src/main/java目录下  
				return null;
            }  
        });  
          
        Template temp = freemarkerConfig.getTemplate("test.xml");  
          
        File targetFile = new File(filePath);  
        Writer out = new OutputStreamWriter(new FileOutputStream(targetFile),"UTF-8");  
          
        //执行模板替换  
        temp.process(result, out);  
        out.flush();  
    }  
    @Test
    public void test() throws TemplateNotFoundException, MalformedTemplateNameException, ParseException, IOException, TemplateException {
    	this.genWordFile("D:\\Eworkspace\\word\\src\\main\\java\\word.doc");
    }
}  