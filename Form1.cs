using System;
using System.Collections;
using System.Collections.Generic;
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
        private string strSource = "source.wav";
        public Form1()
        {
            InitializeComponent();
            for (int i = 0; i < 6; i++)
            {
                if (File.Exists($"{i}_{strSource}"))
                    File.Delete($"{i}_{strSource}");
            }
            PerformRecordStart();
        }


        private void PerformRecordStart()
        {
            var deviceEnum = new MMDeviceEnumerator();
            
            var captureDevices = deviceEnum.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active).ToList();
            var renderDevices = deviceEnum.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToList();
            //devices.ForEach(d => Console.WriteLine(d.FriendlyName, d.AudioClient));
            _Device = CreateWaveInDevice(DataFlow.Capture, captureDevices[1]);
            _Device.StartRecording();
        }

        private void btnRecord_Click(object sender, EventArgs e)
        {
            PerformRecordStart();
        }

        private IWaveIn CreateWaveInDevice(DataFlow flow, MMDevice mmDevice)
        {
            IWaveIn CaptureDevice = null;
            Console.WriteLine($"-------------------{Thread.CurrentThread.ManagedThreadId}");
            if (flow == DataFlow.Capture)
            {
                CaptureDevice = new WasapiCapture(mmDevice) { };
            }

            if (flow == DataFlow.Render)
            {
                var a = mmDevice.AudioEndpointVolume.MasterVolumeLevel;
                mmDevice.AudioEndpointVolume.MasterVolumeLevel = -5;
                var b = mmDevice.AudioEndpointVolume.MasterVolumeLevelScalar;
                //mmDevice.AudioEndpointVolume.MasterVolumeLevelScalar = 1.0f;
                CaptureDevice = new WasapiLoopbackCapture(mmDevice);
            }



            CaptureDevice.DataAvailable += OnDataAvailable;
            CaptureDevice.RecordingStopped += OnRecordingStopped;
            Console.WriteLine($"-------------------Device Created: {CaptureDevice.WaveFormat.Channels}");
            return CaptureDevice;
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


        byte[][] ExtractIndividualChannels(byte[] source, int ChannelCount, int sampleSize)
        {
            int eachChannelSize = source.Length / ChannelCount;
            int FrameCount = eachChannelSize / sizeof(short);

            var ExtractedChannels = new byte[ChannelCount][];
            for (int i = 0; i < ChannelCount; i++)
                ExtractedChannels[i] = new byte[eachChannelSize];

            int BytesInEachFrame = ChannelCount * sampleSize;

            for (int ChannelCounter = 0; ChannelCounter < ChannelCount; ChannelCounter++)
            {
                //Finish One Channel
                int frameCounter = 0;
                for (int byteCounter = sizeof(short)*ChannelCounter; frameCounter < FrameCount; byteCounter += BytesInEachFrame)
                {
                    ExtractedChannels[ChannelCounter][2*frameCounter] = source[byteCounter];
                    ExtractedChannels[ChannelCounter][2 * frameCounter +1] = source[byteCounter+1];
                    frameCounter++;
                }
            }

            return ExtractedChannels;
        }

        private void OnDataAvailable(object sender, WaveInEventArgs e)
        {
#if true
            
            var sourceAudioData = ConvertFloatSamplesIntoShortSamples(e.Buffer, e.BytesRecorded);
            byte[][] res = ExtractIndividualChannels(sourceAudioData, _Device.WaveFormat.Channels, sizeof(short));

            for (int i = 0; i < _Device.WaveFormat.Channels; i++)
            {
                WaveFormat destinationWaveFormat = new WaveFormat(16000, 16, 1);
                byte[] resampledRawData = ResampleRawPcmData(res[i], new WaveFormat(_Device.WaveFormat.SampleRate, sizeof(short) * 8, 1), destinationWaveFormat);


                // Use the converted audio data as needed
                using (var fileStream = new FileStream($"{i}_{strSource}", FileMode.Append, FileAccess.Write, FileShare.None))
                using (var bw = new BinaryWriter(fileStream))
                {
                    bw.Write(res[i], 0, res[i].Length);
                }
            }
#endif
        }

        private IWaveIn _Device;
        private void btnFinish_Click(object sender, EventArgs e)
        {
            _Device?.StopRecording();
        }
    }
}
