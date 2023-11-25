using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Devices.Enumeration;
using Windows.Devices.Midi;

namespace jp.kshoji.unity.midi.uwp
{
    public delegate void OnMidiInputDeviceAttachedHandler(string deviceId);
    public delegate void OnMidiInputDeviceDetachedHandler(string deviceId);
    public delegate void OnMidiOutputDeviceAttachedHandler(string deviceId);
    public delegate void OnMidiOutputDeviceDetachedHandler(string deviceId);

    public delegate void OnMidiNoteOnHandler(string deviceId, byte channel, byte note, byte velocity);
    public delegate void OnMidiNoteOffHandler(string deviceId, byte channel, byte note, byte velocity);
    public delegate void OnMidiPolyphonicKeyPressureHandler(string deviceId, byte channel, byte note, byte velocity);
    public delegate void OnMidiControlChangeHandler(string deviceId, byte channel, byte controller, byte controllerValue);
    public delegate void OnMidiProgramChangeHandler(string deviceId, byte channel, byte program);
    public delegate void OnMidiChannelPressureHandler(string deviceId, byte channel, byte pressure);
    public delegate void OnMidiPitchBendChangeHandler(string deviceId, byte channel, ushort bend);
    public delegate void OnMidiSystemExclusiveHandler(string deviceId, [ReadOnlyArray] byte[] systemExclusive);
    public delegate void OnMidiTimeCodeHandler(string deviceId, byte frameType, byte values);
    public delegate void OnMidiSongPositionPointerHandler(string deviceId, ushort beats);
    public delegate void OnMidiSongSelectHandler(string deviceId, byte song);
    public delegate void OnMidiTuneRequestHandler(string deviceId);
    public delegate void OnMidiTimingClockHandler(string deviceId);
    public delegate void OnMidiStartHandler(string deviceId);
    public delegate void OnMidiContinueHandler(string deviceId);
    public delegate void OnMidiStopHandler(string deviceId);
    public delegate void OnMidiActiveSensingHandler(string deviceId);
    public delegate void OnMidiSystemResetHandler(string deviceId);

    /// <summary>
    /// MIDI Plugin for UWP
    /// </summary>
    public class MidiPlugin
    {
        #region Instantiation
        private static MidiPlugin instance;
        private static readonly object LockObject = new object();

        /// <summary>
        /// Get an instance
        /// </summary>
        public static MidiPlugin Instance
        {
            get
            {
                if (instance != null)
                {
                    return instance;
                }

                lock (LockObject)
                {
                    instance = new MidiPlugin();
                }

                return instance;
            }
        }

        private MidiPlugin()
        {
            inPortDeviceWatcher = DeviceInformation.CreateWatcher(MidiInPort.GetDeviceSelector());
            inPortDeviceWatcher.Added += InPortDeviceAdded;
            inPortDeviceWatcher.Updated += InPortDeviceUpdated;
            inPortDeviceWatcher.Removed += InPortDeviceRemoved;
            inPortDeviceWatcher.Start();

            outPortDeviceWatcher = DeviceInformation.CreateWatcher(MidiOutPort.GetDeviceSelector());
            outPortDeviceWatcher.Added += OutPortDeviceAdded;
            outPortDeviceWatcher.Updated += OutPortDeviceUpdated;
            outPortDeviceWatcher.Removed += OutPortDeviceRemoved;
            outPortDeviceWatcher.Start();
        }

        ~MidiPlugin()
        {
            inPortDeviceWatcher.Added -= InPortDeviceAdded;
            inPortDeviceWatcher.Updated -= InPortDeviceUpdated;
            inPortDeviceWatcher.Removed -= InPortDeviceRemoved;
            inPortDeviceWatcher.Stop();
            inPortDeviceWatcher = null;
            lock (inPorts)
            {
                foreach (var inPort in inPorts)
                {
                    inPort.Value.MessageReceived -= InPortMessageReceived;
                    inPort.Value.Dispose();
                }
                inPorts.Clear();
            }

            outPortDeviceWatcher.Added -= OutPortDeviceAdded;
            outPortDeviceWatcher.Updated -= OutPortDeviceUpdated;
            outPortDeviceWatcher.Removed -= OutPortDeviceRemoved;
            outPortDeviceWatcher.Stop();
            outPortDeviceWatcher = null;
            lock (outPorts)
            {
                foreach (var outPort in outPorts)
                {
                    outPort.Value.Dispose();
                }
                outPorts.Clear();
            }

            lock (deviceInformations)
            {
                deviceInformations.Clear();
            }
        }
        #endregion

        #region MidiDeviceConnection
        private DeviceWatcher inPortDeviceWatcher;
        private DeviceWatcher outPortDeviceWatcher;

        private Dictionary<string, MidiInPort> inPorts = new Dictionary<string, MidiInPort>();
        private Dictionary<string, IMidiOutPort> outPorts = new Dictionary<string, IMidiOutPort>();
        private Dictionary<string, DeviceInformation> deviceInformations = new Dictionary<string, DeviceInformation>();

        public event OnMidiInputDeviceAttachedHandler OnMidiInputDeviceAttached;
        public event OnMidiInputDeviceDetachedHandler OnMidiInputDeviceDetached;
        public event OnMidiOutputDeviceAttachedHandler OnMidiOutputDeviceAttached;
        public event OnMidiOutputDeviceDetachedHandler OnMidiOutputDeviceDetached;

        private async void InPortDeviceAdded(DeviceWatcher deviceWatcher, DeviceInformation deviceInformation)
        {
            var deviceId = deviceInformation.Id;
            if (deviceInformation.Properties.ContainsKey("System.Devices.ContainerId"))
            {
                var containerId = deviceInformation.Properties["System.Devices.ContainerId"];
                var selectorWithContainer = $"System.Devices.ContainerId:=\"{{{containerId}}}\"";
                var serviceInformations = await DeviceInformation.FindAllAsync(selectorWithContainer);
                foreach (var serviceInformation in serviceInformations)
                {
                    if (serviceInformation.Properties.ContainsKey("System.Devices.DeviceInstanceId"))
                    {
                        deviceInformation = serviceInformation;
                        break;
                    }
                }
            }

            lock (deviceInformations)
            {
                if (!deviceInformations.ContainsKey(deviceId))
                {
                    deviceInformations.Add(deviceId, deviceInformation);
                }
            }

            var midiInPort = await MidiInPort.FromIdAsync(deviceId);
            if (midiInPort != null)
            {
                midiInPort.MessageReceived += InPortMessageReceived;
                lock (inPorts)
                {
                    inPorts.Add(deviceId, midiInPort);
                    OnMidiInputDeviceAttached?.Invoke(deviceId);
                }
            }
        }

        private async void InPortDeviceUpdated(DeviceWatcher deviceWatcher, DeviceInformationUpdate deviceInformation)
        {
            var midiInPort = await MidiInPort.FromIdAsync(deviceInformation.Id);
            lock (inPorts)
            {
                if (!inPorts.ContainsKey(deviceInformation.Id))
                {
                    inPorts.Add(deviceInformation.Id, midiInPort);
                    OnMidiInputDeviceAttached?.Invoke(deviceInformation.Id);
                }
            }
        }

        private void InPortDeviceRemoved(DeviceWatcher deviceWatcher, DeviceInformationUpdate deviceInformation)
        {
            lock (deviceInformations)
            {
                if (deviceInformations.ContainsKey(deviceInformation.Id))
                {
                    deviceInformations.Remove(deviceInformation.Id);
                }
            }

            lock (inPorts)
            {
                if (inPorts.ContainsKey(deviceInformation.Id))
                {
                    inPorts[deviceInformation.Id].MessageReceived -= InPortMessageReceived;
                    inPorts[deviceInformation.Id].Dispose();
                    inPorts.Remove(deviceInformation.Id);
                    OnMidiInputDeviceDetached?.Invoke(deviceInformation.Id);
                }
            }
        }

        private async void OutPortDeviceAdded(DeviceWatcher deviceWatcher, DeviceInformation deviceInformation)
        {
            var deviceId = deviceInformation.Id;
            if (deviceInformation.Properties.ContainsKey("System.Devices.ContainerId"))
            {
                var containerId = deviceInformation.Properties["System.Devices.ContainerId"];
                var selectorWithContainer = $"System.Devices.ContainerId:=\"{{{containerId}}}\"";
                var serviceInformations = await DeviceInformation.FindAllAsync(selectorWithContainer);
                foreach (var serviceInformation in serviceInformations)
                {
                    if (serviceInformation.Properties.ContainsKey("System.Devices.DeviceInstanceId"))
                    {
                        deviceInformation = serviceInformation;
                        break;
                    }
                }
            }

            lock (deviceInformations)
            {
                if (!deviceInformations.ContainsKey(deviceId))
                {
                    deviceInformations.Add(deviceId, deviceInformation);
                }
            }

            var midiOutPort = await MidiOutPort.FromIdAsync(deviceId);
            if (midiOutPort != null)
            {
                lock (outPorts)
                {
                    outPorts.Add(deviceId, midiOutPort);
                    OnMidiOutputDeviceAttached?.Invoke(deviceId);
                }
            }
        }

        private async void OutPortDeviceUpdated(DeviceWatcher deviceWatcher, DeviceInformationUpdate deviceInformation)
        {
            var midiOutPort = await MidiOutPort.FromIdAsync(deviceInformation.Id);

            lock (outPorts)
            {
                if (!outPorts.ContainsKey(deviceInformation.Id))
                {
                    outPorts.Add(deviceInformation.Id, midiOutPort);
                    OnMidiOutputDeviceAttached?.Invoke(deviceInformation.Id);
                }
            }
        }

        private void OutPortDeviceRemoved(DeviceWatcher deviceWatcher, DeviceInformationUpdate deviceInformation)
        {
            lock (deviceInformations)
            {
                if (deviceInformations.ContainsKey(deviceInformation.Id))
                {
                    deviceInformations.Remove(deviceInformation.Id);
                }
            }

            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceInformation.Id))
                {
                    outPorts[deviceInformation.Id].Dispose();
                    outPorts.Remove(deviceInformation.Id);
                    OnMidiOutputDeviceDetached?.Invoke(deviceInformation.Id);
                }
            }
        }

        /// <summary>
        /// Get the device name from specified device ID.
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <returns>the device name, empty if device not connected.</returns>
        public string GetDeviceName(string deviceId)
        {
            lock (deviceInformations)
            {
                if (deviceInformations.ContainsKey(deviceId))
                {
                    return deviceInformations[deviceId].Name;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Get the device vendor id from specified device ID.
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <returns>the device name, empty if device not connected.</returns>
        public string GetVendorId(string deviceId)
        {
            lock (deviceInformations)
            {
                if (deviceInformations.ContainsKey(deviceId))
                {
                    if (deviceInformations[deviceId].Properties.TryGetValue("System.Devices.DeviceInstanceId", out var deviceInstanceId))
                    {
                        var deviceInstanceIdString = deviceInstanceId as string;
                        if (deviceInstanceIdString == null)
                        {
                            return string.Empty;
                        }

                        // USB\VID_XXXX&PID_XXXX\xxxxxxxx
                        var splitted = deviceInstanceIdString.Split('\\');
                        if (splitted.Length > 1)
                        {
                            var vidPidSplitted = splitted[1].Split('&');
                            if (vidPidSplitted.Length > 1)
                            {
                                return vidPidSplitted[0];
                            }
                        }
                    }
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Get the device product id from specified device ID.
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <returns>the device name, empty if device not connected.</returns>
        public string GetProductId(string deviceId)
        {
            lock (deviceInformations)
            {
                if (deviceInformations.ContainsKey(deviceId))
                {
                    if (deviceInformations[deviceId].Properties.TryGetValue("System.Devices.DeviceInstanceId", out var deviceInstanceId))
                    {
                        var deviceInstanceIdString = deviceInstanceId as string;
                        if (deviceInstanceIdString == null)
                        {
                            return string.Empty;
                        }

                        // USB\VID_XXXX&PID_XXXX\xxxxxxxx
                        var splitted = deviceInstanceIdString.Split('\\');
                        if (splitted.Length > 1)
                        {
                            var vidPidSplitted = splitted[1].Split('&');
                            if (vidPidSplitted.Length > 1)
                            {
                                return vidPidSplitted[1];
                            }
                        }
                    }
                }
            }

            return string.Empty;
        }
        #endregion

        #region MidiEventReceiving
        public event OnMidiNoteOnHandler OnMidiNoteOn;
        public event OnMidiNoteOffHandler OnMidiNoteOff;
        public event OnMidiPolyphonicKeyPressureHandler OnMidiPolyphonicKeyPressure;
        public event OnMidiControlChangeHandler OnMidiControlChange;
        public event OnMidiProgramChangeHandler OnMidiProgramChange;
        public event OnMidiChannelPressureHandler OnMidiChannelPressure;
        public event OnMidiPitchBendChangeHandler OnMidiPitchBendChange;
        public event OnMidiSystemExclusiveHandler OnMidiSystemExclusive;
        public event OnMidiTimeCodeHandler OnMidiTimeCode;
        public event OnMidiSongPositionPointerHandler OnMidiSongPositionPointer;
        public event OnMidiSongSelectHandler OnMidiSongSelect;
        public event OnMidiTuneRequestHandler OnMidiTuneRequest;
        public event OnMidiTimingClockHandler OnMidiTimingClock;
        public event OnMidiStartHandler OnMidiStart;
        public event OnMidiContinueHandler OnMidiContinue;
        public event OnMidiStopHandler OnMidiStop;
        public event OnMidiActiveSensingHandler OnMidiActiveSensing;
        public event OnMidiSystemResetHandler OnMidiSystemReset;

        private void InPortMessageReceived(MidiInPort sender, MidiMessageReceivedEventArgs args)
        {
            IMidiMessage receivedMidiMessage = args.Message;

            Debug.WriteLine(receivedMidiMessage.Timestamp.ToString());

            switch (receivedMidiMessage.Type)
            {
                case MidiMessageType.NoteOff:
                    {
                        var noteOff = (MidiNoteOffMessage)receivedMidiMessage;
                        OnMidiNoteOff?.Invoke(sender.DeviceId, noteOff.Channel, noteOff.Note, noteOff.Velocity);
                    }
                    break;
                case MidiMessageType.NoteOn:
                    {
                        var noteOn = (MidiNoteOnMessage)receivedMidiMessage;
                        OnMidiNoteOn?.Invoke(sender.DeviceId, noteOn.Channel, noteOn.Note, noteOn.Velocity);
                    }
                    break;
                case MidiMessageType.PolyphonicKeyPressure:
                    {
                        var polyphonincKeyPressure = (MidiPolyphonicKeyPressureMessage)receivedMidiMessage;
                        OnMidiPolyphonicKeyPressure?.Invoke(sender.DeviceId, polyphonincKeyPressure.Channel, polyphonincKeyPressure.Note, polyphonincKeyPressure.Pressure);
                    }
                    break;
                case MidiMessageType.ControlChange:
                    {
                        var controlChange = (MidiControlChangeMessage)receivedMidiMessage;
                        OnMidiControlChange?.Invoke(sender.DeviceId, controlChange.Channel, controlChange.Controller, controlChange.ControlValue);
                    }
                    break;
                case MidiMessageType.ProgramChange:
                    {
                        var programChange = (MidiProgramChangeMessage)receivedMidiMessage;
                        OnMidiProgramChange?.Invoke(sender.DeviceId, programChange.Channel, programChange.Program);
                    }
                    break;
                case MidiMessageType.ChannelPressure:
                    {
                        var channelPressure = (MidiChannelPressureMessage)receivedMidiMessage;
                        OnMidiChannelPressure?.Invoke(sender.DeviceId, channelPressure.Channel, channelPressure.Pressure);
                    }
                    break;
                case MidiMessageType.PitchBendChange:
                    {
                        var pitchBendChange = (MidiPitchBendChangeMessage)receivedMidiMessage;
                        OnMidiPitchBendChange?.Invoke(sender.DeviceId, pitchBendChange.Channel, pitchBendChange.Bend);
                    }
                    break;
                case MidiMessageType.SystemExclusive:
                    {
                        var systemExclusive = (MidiSystemExclusiveMessage)receivedMidiMessage;
                        OnMidiSystemExclusive?.Invoke(sender.DeviceId, systemExclusive.RawData.ToArray());
                    }
                    break;
                case MidiMessageType.MidiTimeCode:
                    {
                        var midiTimeCode = (MidiTimeCodeMessage)receivedMidiMessage;
                        OnMidiTimeCode?.Invoke(sender.DeviceId, midiTimeCode.FrameType, midiTimeCode.Values);
                    }
                    break;
                case MidiMessageType.SongPositionPointer:
                    {
                        var songPositionPointer = (MidiSongPositionPointerMessage)receivedMidiMessage;
                        OnMidiSongPositionPointer?.Invoke(sender.DeviceId, songPositionPointer.Beats);
                    }
                    break;
                case MidiMessageType.SongSelect:
                    {
                        var songSelect = (MidiSongSelectMessage)receivedMidiMessage;
                        OnMidiSongSelect?.Invoke(sender.DeviceId, songSelect.Song);
                    }
                    break;
                case MidiMessageType.TuneRequest:
                    OnMidiTuneRequest?.Invoke(sender.DeviceId);
                    break;
                case MidiMessageType.TimingClock:
                    OnMidiTimingClock?.Invoke(sender.DeviceId);
                    break;
                case MidiMessageType.Start:
                    OnMidiStart?.Invoke(sender.DeviceId);
                    break;
                case MidiMessageType.Continue:
                    OnMidiContinue?.Invoke(sender.DeviceId);
                    break;
                case MidiMessageType.Stop:
                    OnMidiStop?.Invoke(sender.DeviceId);
                    break;
                case MidiMessageType.ActiveSensing:
                    OnMidiActiveSensing?.Invoke(sender.DeviceId);
                    break;
                case MidiMessageType.SystemReset:
                    OnMidiSystemReset?.Invoke(sender.DeviceId);
                    break;

                case MidiMessageType.EndSystemExclusive:
                case MidiMessageType.None:
                default:
                    break;
            }
        }
        #endregion

        #region MidiEventSending
        /// <summary>
        /// Send a NoteOn message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <param name="channel">the channel(0-15)</param>
        /// <param name="note">the note number(0-127)</param>
        /// <param name="velocity">the velocity(0-127)</param>
        public void SendMidiNoteOn(string deviceId, byte channel, byte note, byte velocity)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiNoteOnMessage(channel, note, velocity));
                }
            }
        }

        /// <summary>
        /// Send a NoteOff message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <param name="channel">the channel(0-15)</param>
        /// <param name="note">the note number(0-127)</param>
        /// <param name="velocity">the velocity(0-127)</param>
        public void SendMidiNoteOff(string deviceId, byte channel, byte note, byte velocity)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiNoteOffMessage(channel, note, velocity));
                }
            }
        }

        /// <summary>
        /// Send a PolyphonicKeyPressure message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <param name="channel">the channel(0-15)</param>
        /// <param name="note">the note number(0-127)</param>
        /// <param name="velocity">the velocity(0-127)</param>
        public void SendMidiPolyphonicKeyPressure(string deviceId, byte channel, byte note, byte velocity)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiPolyphonicKeyPressureMessage(channel, note, velocity));
                }
            }
        }

        /// <summary>
        /// Send a ControlChange message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <param name="channel">the channel(0-15)</param>
        /// <param name="controller">the controller(0-127)</param>
        /// <param name="controllerValue">the value(0-127)</param>
        public void SendMidiControlChange(string deviceId, byte channel, byte controller, byte controllerValue)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiControlChangeMessage(channel, controller, controllerValue));
                }
            }
        }

        /// <summary>
        /// Send a ProgramChange message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <param name="channel">the channel(0-15)</param>
        /// <param name="program">the program(0-127)</param>
        public void SendMidiProgramChange(string deviceId, byte channel, byte program)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiProgramChangeMessage(channel, program));
                }
            }
        }

        /// <summary>
        /// Send a ChannelPressure message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <param name="channel">the channel(0-15)</param>
        /// <param name="pressure">the pressure(0-127)</param>
        public void SendMidiChannelPressure(string deviceId, byte channel, byte pressure)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiChannelPressureMessage(channel, pressure));
                }
            }
        }

        /// <summary>
        /// Send a PitchBendChange message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <param name="channel">the channel(0-15)</param>
        /// <param name="bend">the pitch bend value(0-16383)</param>
        public void SendMidiPitchBendChange(string deviceId, byte channel, ushort bend)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiPitchBendChangeMessage(channel, bend));
                }
            }
        }

        /// <summary>
        /// Send a SystemExclusive message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <param name="systemExclusive">the system exclusive data</param>
        public void SendMidiSystemExclusive(string deviceId, [ReadOnlyArray] byte[] systemExclusive)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiSystemExclusiveMessage(systemExclusive.AsBuffer()));
                }
            }
        }

        /// <summary>
        /// Send a TimeCode message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <param name="frameType">the frame type(0-7)</param>
        /// <param name="values">the time code(0-15)</param>
        public void SendMidiTimeCode(string deviceId, byte frameType, byte values)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiTimeCodeMessage(frameType, values));
                }
            }
        }

        /// <summary>
        /// Send a SongPositionPointer message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <param name="beats">the song position pointer(0-16383)</param>
        public void SendMidiSongPositionPointer(string deviceId, ushort beats)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiSongPositionPointerMessage(beats));
                }
            }
        }

        /// <summary>
        /// Send a SongSelect message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        /// <param name="song"></param>
        public void SendMidiSongSelect(string deviceId, byte song)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiSongSelectMessage(song));
                }
            }
        }

        /// <summary>
        /// Send a TuneRequest message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        public void SendMidiTuneRequest(string deviceId)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiTuneRequestMessage());
                }
            }
        }

        /// <summary>
        /// Send a TimingClock message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        public void SendMidiTimingClock(string deviceId)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiTimingClockMessage());
                }
            }
        }

        /// <summary>
        /// Send a Start message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        public void SendMidiStart(string deviceId)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiStartMessage());
                }
            }
        }

        /// <summary>
        /// Send a Continue message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        public void SendMidiContinue(string deviceId)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiContinueMessage());
                }
            }
        }

        /// <summary>
        /// Send a Stop message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        public void SendMidiStop(string deviceId)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiStopMessage());
                }
            }
        }

        /// <summary>
        /// Send an ActiveSensing message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        public void SendMidiActiveSensing(string deviceId)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiActiveSensingMessage());
                }
            }
        }

        /// <summary>
        /// Send a SystemReset message
        /// </summary>
        /// <param name="deviceId">the device ID</param>
        public void SendMidiSystemReset(string deviceId)
        {
            lock (outPorts)
            {
                if (outPorts.ContainsKey(deviceId))
                {
                    outPorts[deviceId].SendMessage(new MidiSystemResetMessage());
                }
            }
        }
        #endregion
    }
}
