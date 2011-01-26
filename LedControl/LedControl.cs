/*
 * Created by SharpDevelop.
 * User: oferfrid
 * Date: 19/03/2009
 * Time: 13:53
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

//In order to create the .tlb file use the tlbexp assemblyName [/out:file] command

using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Net.Sockets;
using System.Runtime.InteropServices;

namespace CoolLedControl
{
	//F96CF5F3-5651-40ED-AFAD-B18AC331AD22
	[Guid("1B9618B4-BD4F-4CF5-82C8-830885F5265F")]
	
	
	
	public interface ILedControl
	{
		[DispId(1)]
		string openLed(int ind);
		[DispId(2)]
		string openLed(int ind,int msDelay);
		[DispId(3)]
		string closeLed(int ind);
		[DispId(4)]
		string changeIntensity(int ind, int intensity);
		[DispId(5)]
		void openSock();
		[DispId(6)]
		void closeSock();
	}
	
	


	/// <summary>
	/// Description of LedControl.
	/// DLL to open, close and change the intesntity of coolLed.
	/// There are 4 channels with different wavelength A-D.
	/// The intensity is 0-100.
	///</summary>
	
	
	//D718A292-71A3-4E69-AC2D-FFAE9990BC2B
	[Guid("1EF64472-ECAB-419A-B53B-9818194BD07D"),
	 ClassInterface(ClassInterfaceType.None)]
	public class LedControl:ILedControl
	{
		TcpClient tcpclnt;
		public LedControl()
		{
			
		}
		~LedControl()
		{
			tcpclnt.Close();
		}
		
		public string openLed(int ind)
		{
			string str ="C"+int2channel(ind)+"N";
			return sendString(str);
		}
		
		public string openLed(int ind,int msDelay)
		{
			string str ="C"+int2channel(ind)+"N";
			string res = sendString(str);
			System.Threading.Thread.Sleep(msDelay);
			return res;
		}
		
		public string closeLed(int ind){
			string str ="C"+int2channel(ind)+"F";
			return sendString(str);
		}
		
		public string changeIntensity(int ind, int intensity){
			string str = "C"+int2channel(ind)+"I"+intensity;
			return sendString(str);
		}
		public void openSock(){
			tcpclnt = new TcpClient();
			tcpclnt.Connect("192.168.0.252",18259); // use the ipaddress as in the server program
			//should the IP addres and port number be used as parameters????
		}
		
		public void closeSock()
		{
			
			tcpclnt.Close();
		}
		//channels are A,B,C,D. and their relative ints are 1-4
		private char int2channel(int ind){
			int asciiCode = 64+ind;
			return (char)asciiCode;
		}
		
		private string sendString(string str){
			Stream stm = tcpclnt.GetStream();
			
			ASCIIEncoding asen= new ASCIIEncoding();
			byte[] ba=asen.GetBytes(str+ "\n");
			stm.Write(ba,0,ba.Length);
			
			//receiving response???
			System.Threading.Thread.Sleep(10);
			byte[] bb=new byte[100];
			int k=stm.Read(bb,0,100);
			string s = System.Text.Encoding.UTF8.GetString(bb,0,k);
			
			return s;
			
			

		}
	}
}




