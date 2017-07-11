package com.gavs;

import java.io.IOException;

import java.util.Date;
import java.util.Timer;
import java.util.TimerTask;

/**
 * @author sreenivasulu.s
 *
 */

public class DefaultDemoTimerTask extends TimerTask {
	private final static long ONCE_PER_DAY = 1000 * 60 * 60 * 24;

	// private final static int ONE_DAY = 1;
	private final static int SEVEN_AM = 00;
	private final static int ZERO_MINUTES = 02;

	@Override
	public void run() {
		{
			try {

				TicketsImport.uploadSolrData();
				Opportunities.uploadSolrData();
				MonitoringProblems.uploadSolrData();
				CIAssetsImport.uploadSolrData();
				CIApplications.uploadSolrData();
				DefaultTechnicians.uploadSolrData();

				MonitoringCPU.uploadSolrData();
				CPUForecast.uploadSolrData();
				MonitoringMemory.uploadSolrData();
				MemoryForecast.uploadSolrData();
				MonitoringDiskOpen.uploadSolrData();
				OpenDiskPrediction.uploadSolrData();
				
				BMSMonitoringDisk.uploadSolrData();
				BMSMonitoringCPU.uploadSolrData();
				BSMMonitoringMemory.uploadSolrData();

				MonitoringCPU2.uploadSolrData();
				CPUForecast2.uploadSolrData();
				MonitoringCPU3.uploadSolrData();
				CPUForecast3.uploadSolrData();
				CPUPrediction.uploadSolrData();

				MonitoringMemoryClosed.uploadSolrData();
				MemoryForecastClosed.uploadSolrData();
				MonitoringMemoryClosed2.uploadSolrData();
				MemoryForecastClosed2.uploadSolrData();
				MemoryPrediction.uploadSolrData();

				MonitoringDiskClosed.uploadSolrData();
				ClosedDiskPrediction.uploadSolrData();
				MonitoringDiskClosed2.uploadSolrData();
				ClosedDiskPrediction2.uploadSolrData();

				MonitoringMemoryDeclined.uploadSolrData();
				MemoryForecastDeclined.uploadSolrData();

				MonitoringInProgressCPU.uploadSolrData();
				CPUInProgressForecast.uploadSolrData();

				MonitoringMemoryInProgress.uploadSolrData();
				MemoryForecastInProgress.uploadSolrData();

				MonitoringAlerts.uploadSolrData();
				MonitoringPing.uploadSolrData();

				System.out.println("successfuly loaded in to Solr--- > " + new Date());
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	private static Date getTomorrowMorning7AM() {

		Date date2am = new java.util.Date();
		date2am.setHours(SEVEN_AM);
		date2am.setMinutes(ZERO_MINUTES);

		return date2am;
	}

	public static void startTask() {
		DefaultDemoTimerTask task = new DefaultDemoTimerTask();
		Timer timer = new Timer();
		timer.schedule(task, getTomorrowMorning7AM(), ONCE_PER_DAY);

	}

	public static void main(String args[]) {
		startTask();

	}

}
