package com.sage.live.migrationservice.plugin.fec;

import javax.swing.filechooser.FileFilter;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import org.pf4j.Extension;
import org.pf4j.Plugin;
import org.pf4j.PluginWrapper;

import com.sage.live.migrationservice.plugin.PluginInfo;
import com.sage.live.migrationservice.plugin.Utilities;

import com.sage.live.migrationservice.converter.*;

public class FECPlugin extends Plugin {
	// implements com.sage.live.migrationservice.plugin.MigrationPlugin {

	public FECPlugin(PluginWrapper wrapper) {
		super(wrapper);
	}

	@Override
	public void start() {
		System.out.println("FEC Plugin.start()");
	}

	@Override
	public void stop() {
		System.out.println("FEC Plugin.stop()");
	}

	@Extension
	public static class FECConverter implements com.sage.live.migrationservice.plugin.MigrationPlugin {
		public boolean convertSourceToTarget(String sourceFileName, String targetFileName) {
			Boolean result = false;
			try {
				
				File logFile = new File("./logs/convert.log");
				
				BufferedWriter writer = new BufferedWriter(new FileWriter(logFile));
				//
				 FEC2xlsx x = new FEC2xlsx();
				 result = x.Convert(sourceFileName,targetFileName,(aLogType,aMessage) -> {
					try {
						writer.write(aLogType.toString() + " " + aMessage);
						writer.newLine();
					} catch (IOException e) {
						e.printStackTrace();
					}
				});				
				writer.close();
				
			} catch (Exception ex) {
				return false;
			}
			return result;
		}

		public String getVersion() {
			Package packageObject = FECConverter.class.getPackage();
			return packageObject.getSpecificationVersion();
		}

		public List<PluginInfo> getPluginInfo() {
			ArrayList<PluginInfo> infos = new ArrayList<PluginInfo>();

			PluginInfo info = new PluginInfo(this.getClass().getCanonicalName(), "FEC File");
			info.addFileFilter(new SampleFileFilter());
			info.setPluginInstance(this);

			infos.add(info);

			return infos;
		}

		private class SampleFileFilter extends FileFilter {
			public String getDescription() {
				return "FEC files";
			}

			@Override
			public boolean accept(File f) {
				if (f.isDirectory())
					return true;
				String extension = Utilities.getFileExtension(f);
				if (extension != null) {
					if (extension.toLowerCase().equals("txt")) {
						return true;
					}
				}
				return false;
			}
		}
	}
}
