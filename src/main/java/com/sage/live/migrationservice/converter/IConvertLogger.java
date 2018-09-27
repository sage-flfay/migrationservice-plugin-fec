package com.sage.live.migrationservice.converter;

@FunctionalInterface
public interface IConvertLogger <TLogType , TString > {
	
	public void doLog(LogType aLogType, String aMessage);	
} 
