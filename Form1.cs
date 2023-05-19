using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting.Channels;
using System.Threading;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.Compression;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrayNotify;

namespace WasapiLoopback
{
    public partial class Form1 : Form
    {
        private string strOut = "out.wav";
        private string strSource = "source.wav";
        public Form1()
        {
            InitializeComponent();
            if (File.Exists(strSource))
                File.Delete(strSource);
            if (File.Exists(strOut))
                File.Delete(strOut);
            PerformRecordStart();
        }


        private void PerformRecordStart()
        {
            captureDevice = CreateWaveInDevice();

            writer = new WaveFileWriter(strOut, captureDevice.WaveFormat);
            captureDevice.StartRecording();
        }

        private void btnRecord_Click(object sender, EventArgs e)
        {
            PerformRecordStart();
        }

        private IWaveIn newWaveIn;
        private IWaveIn CreateWaveInDevice()
        {
            Console.WriteLine($"-------------------{Thread.CurrentThread.ManagedThreadId}");
            var deviceEnum = new MMDeviceEnumerator();
            var renderDevices = deviceEnum.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToList();
            renderDevices.ForEach(d=>Console.WriteLine(d.FriendlyName, d.AudioClient));

            var device = (MMDevice)renderDevices[2];



            newWaveIn = new WasapiLoopbackCapture(device);

            var rr = newWaveIn.WaveFormat.Channels;


            // Both WASAPI and WaveIn support Sample Rate conversion!
            //var sampleRate = (int)16000;
            //var channels = 1;
            //newWaveIn.WaveFormat = new WaveFormat(sampleRate, channels);
            //newWaveIn.WaveFormat = WaveFormat.CreateCustomFormat(WaveFormatEncoding.Pcm, 16000, 6, 32000*6, 2*6,2);

            newWaveIn.DataAvailable += OnDataAvailable;
            newWaveIn.RecordingStopped += OnRecordingStopped;
            return newWaveIn;
        }

        private void OnRecordingStopped(object sender, StoppedEventArgs e)
        {
            Console.WriteLine("OnRecordingStopped");
            writer?.Dispose();
            writer = null;
        }

        private WaveFileWriter writer;




        private byte[] ConvertFloatSamplesIntoShortSamples(byte[] samples, int bytesCount)
        {
            int samplesCount = bytesCount / 4;
            var pcm = new byte[samplesCount * 2];
            int byteIndex = 0,
                pcmIndex = 0;

            while (byteIndex < bytesCount)
            {
                float floatValue = BitConverter.ToSingle(samples, byteIndex);

                var outsample = (short)(floatValue * short.MaxValue);
                pcm[pcmIndex] = (byte)(outsample & 0xff);
                pcm[pcmIndex + 1] = (byte)((outsample >> 8) & 0xff);

                byteIndex += 4;
                pcmIndex += 2;
            }

            return pcm;
        }


        byte[] ResampleRawPcmData(byte[] sourceAudioData, WaveFormat sourceWaveFormat, WaveFormat destinationWaveFormat)
        {
            RawSourceWaveStream sourceStream = new RawSourceWaveStream(new MemoryStream(sourceAudioData), sourceWaveFormat);

            using (MemoryStream outputStream = new MemoryStream())
                using (WaveFormatConversionStream conversionStream = new WaveFormatConversionStream(destinationWaveFormat, sourceStream))
            using (RawSourceWaveStream rawDestinationStream =
                   new RawSourceWaveStream(conversionStream, destinationWaveFormat))
            {
                byte[] convertedAudioData = new byte[rawDestinationStream.Length];
                rawDestinationStream.Read(convertedAudioData, 0, convertedAudioData.Length);
                return convertedAudioData;
            }
        }

        private void OnDataAvailable(object sender, WaveInEventArgs e)
        {
            //Console.WriteLine($"OnDataAvailable: {e.BytesRecorded} {Thread.CurrentThread.ManagedThreadId}");

            var channelCount = newWaveIn.WaveFormat.Channels;
#if true
            WaveFormat sourceWaveFormat = new WaveFormat(48000, 16, 2);

            WaveFormat destinationWaveFormat = new WaveFormat(16000, 16, 2);


            var sourceAudioData = ConvertFloatSamplesIntoShortSamples(e.Buffer, e.BytesRecorded);
            byte[] resampledRawData = ResampleRawPcmData(sourceAudioData, sourceWaveFormat, destinationWaveFormat);

            // Use the converted audio data as needed
            using (var fileStream = new FileStream(strSource, FileMode.Append,
                       FileAccess.Write, FileShare.None))
            using (var bw = new BinaryWriter(fileStream))
            {
                bw.Write(resampledRawData, 0, resampledRawData.Length);
            }

#endif


            writer.Write(e.Buffer, 0, e.BytesRecorded);
        }

        private IWaveIn captureDevice;
        private void btnFinish_Click(object sender, EventArgs e)
        {
            captureDevice?.StopRecording();
        }
    }
}



























