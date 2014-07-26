import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.Map;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

public class GetWord {
 Configuration configuration = null;
 public GetWord() {
  configuration = new Configuration();
  configuration.setDefaultEncoding("utf-8");
 }
 /**
  *@param template is word template who is stored in cn/edu/bupt/template and whose suffix is .ftl
  *@param docfile is a doc file name and whose path like "E:/test.doc"
  *@param data is what should store in word and whose key must be consistent with word's key
  *
  *@return null
  *@author yongsheng
  */
 public void createDoc(String template, String docFile, Map<String, String> data) {
      configuration.setClassForTemplateLoading(this.getClass(), "/cn/edu/bupt/template"); //模板放在此包下面
      Template t = null;
      try {
       // test.ftl为要装载的模板
       t = configuration.getTemplate(template);
       t.setEncoding("utf-8");
      } catch (IOException e) {
       e.printStackTrace();
      }
      // 输出文档路径及名称
      File outFile = new File(docFile);
      Writer out = null;
      try {
       out = new BufferedWriter(new OutputStreamWriter(
        new FileOutputStream(outFile), "utf-8"));
      } catch (Exception e1) {
       e1.printStackTrace();
      }
      try {
       t.process(data, out);
       out.close();
      } catch (TemplateException e) {
       e.printStackTrace();
      } catch (IOException e) {
       e.printStackTrace();
      }
 }
}