import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import com.opensymphony.xwork2.ActionContext;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

/**
 *@author yongsheng
 *@version 1.0
 */

public class Export {

     /**
      *@author yongsheng
      *@param data who is what data we should insert into word, whoes key must bee along
      * with word
      *@param template is what template your need. like "test.ftl"
      *@param where is the palce template is stored. like "cn/edu/bupt/template"
      *
      */
      public void exportWord(Map<String, String> data, String template, String where){
      Configuration configuration = new Configuration();
      configuration.setDefaultEncoding("utf-8");
      configuration.setClassForTemplateLoading(getClass(), where); 
      Template t = null;
      try {
       t = configuration.getTemplate(template);// 载入模版文件（把要导出的word模版另存为word
              // xml就可以了）
       t.setEncoding("utf-8");
      } catch (IOException e) {
       e.printStackTrace();
      }
      String fileName = template.substring(0, template.length()-3) + "doc";
      ActionContext ctx = ActionContext.getContext();
      HttpServletResponse response = (HttpServletResponse) ctx
        .get("com.opensymphony.xwork2.dispatcher.HttpServletResponse");
      response.setContentType("application/msword");
      try {
       response.addHeader("Content-Disposition", "attachment; filename="
         + URLEncoder.encode(fileName, "UTF-8"));
       response.setCharacterEncoding("utf-8");
       PrintWriter out = response.getWriter();
       t.process(data, out);
       out.close();
      } catch (UnsupportedEncodingException e) {
       e.printStackTrace();
      } catch (IOException e) {
       e.printStackTrace();
      } catch (TemplateException e) {
       e.printStackTrace();
      }
     }
   }
