/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package localization;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Vector;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 *
 * @author wei7771
 */
public class EsriPslLpu2Excel {
 
    public static void export(String filepath){
        String folder = filepath.substring(0,filepath.lastIndexOf("\\"));
        Vector<String> zFile;
        if(filepath.endsWith(".zip")){
            zFile = readzipfile(filepath);
            for(String s : zFile){
                if(s.endsWith(".lpu")){
                    export(s);
                }
            }
        }
        else if(!filepath.endsWith(".lpu")){
            return;
        }
        else{
            String os = System.getProperty("os.arch");
            String passoloPath = "";
            if(os.contains("x86")){
                passoloPath = "\"C:\\Program Files\\SDL Passolo 2011\\pslcmd.exe\"";
            }
            else{
		passoloPath = "\"C:\\Program Files (x86)\\SDL Passolo 2011\\pslcmd.exe\"";
            }
            
            String logfile = folder + "\\" + filepath.substring(filepath.lastIndexOf("\\")+1,filepath.lastIndexOf(".")) + ".log";
	 	try{
                    File log = new File(logfile);
                    if(!log.exists()){
		 	log.createNewFile();
                    }
	 	}catch(Exception e){
             }
            
             int exitVal = 0;
             do{
                 String cmd2 = passoloPath + " /openproject:" + filepath + " /runmacro=PslLpu2Excel.bas";
                 try{
                     String osName = System.getProperty("os.name");
                     String[] cmd = new String[3];
                     if(osName.equals("Windows NT") || (osName.equals("Windows 7"))){
                         cmd[0] = "cmd.exe";
                         cmd[1] = "/C";
                         cmd[2] = cmd2 + " >> " + logfile;
                     }
                     else if(osName.equals("Windows 95")){
                         cmd[0] = "command.com";
                         cmd[1] = "/C";
                         cmd[2] = cmd2 + " >> " + logfile;
                     }
                     Runtime rt = Runtime.getRuntime();
                     Process proc = rt.exec(cmd);
                     StreamGobbler errorGobbler = new StreamGobbler(proc.getErrorStream(),"ERROR");
                     StreamGobbler outputGobbler = new StreamGobbler(proc.getInputStream(),"OUTPUT");
                     errorGobbler.start();
                     outputGobbler.start();
                     exitVal = proc.waitFor();
                 
                 }catch(Exception e){
                     
                 }
             }while(exitVal != 0);
            
            
            
        }
    }

     public static Vector<String> readzipfile(String filepath){
        Vector<String> v = new Vector<String>();
        byte[] buffer = new byte[1024];
        String outputFolder = filepath.substring(0,filepath.lastIndexOf("."));
        try{
            File folder = new File(outputFolder);
            if(!folder.exists()){
		folder.mkdir();
	  }
			
                ZipInputStream zis = new ZipInputStream(new FileInputStream(filepath));
		ZipEntry ze = zis.getNextEntry();
                while(ze != null){
                    String fileName = ze.getName();
                    File subFolder = new File(outputFolder + "\\" + fileName.substring(0,fileName.lastIndexOf(".")));
                    if(!subFolder.exists()){
                        subFolder.mkdir();
                    }
		    File newFile = new File(outputFolder + "\\" + fileName.substring(0,fileName.lastIndexOf(".")) + "\\" + fileName);
		    v.addElement(newFile.getAbsolutePath());
		    FileOutputStream fos = new FileOutputStream(newFile);
		    int len;
		    while((len = zis.read(buffer)) > 0){
			fos.write(buffer, 0, len);
		     }
                fos.close();
                ze = zis.getNextEntry();
              }	 
	    zis.closeEntry();
	    zis.close();  
        }catch(Exception e){
            
        }
        return v;
    } 
    
}
