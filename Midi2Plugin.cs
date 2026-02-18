using Microsoft.Windows.Devices.Midi2;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;

namespace jp.kshoji.unity.midi.uwp
{
    public delegate void OnMidi2InputDeviceAttachedHandler(string deviceId);
    public delegate void OnMidi2InputDeviceDetachedHandler(string deviceId);
    public delegate void OnMidi2OutputDeviceAttachedHandler(string deviceId);
    public delegate void OnMidi2OutputDeviceDetachedHandler(string deviceId);
    public delegate void OnMidi2UmpHandler(string deviceId, [ReadOnlyArray] uint[] umps);

    /// <summary>
    /// MIDI 2.0 Plugin for UWP
    /// </summary>
    public class Midi2Plugin
    {
        private static Midi2Plugin instance;
        private static readonly object LockObject = new object();

        public event OnMidi2InputDeviceAttachedHandler OnMidi2InputDeviceAttached;
        public event OnMidi2InputDeviceDetachedHandler OnMidi2InputDeviceDetached;
        public event OnMidi2OutputDeviceAttachedHandler OnMidi2OutputDeviceAttached;
        public event OnMidi2OutputDeviceDetachedHandler OnMidi2OutputDeviceDetached;
        public event OnMidi2UmpHandler OnMidi2UmpReceived;

        private MidiEndpointDeviceWatcher midi2DeviceWatcher;
        private MidiSession midi2Session;
        private Dictionary<string, MidiEndpointConnection> midi2Connections = new Dictionary<string, MidiEndpointConnection>();

        /// <summary>
        /// Get an instance
        /// </summary>
        public static Midi2Plugin Instance
        {
            get
            {
                if (instance != null)
                {
                    return instance;
                }

                lock (LockObject)
                {
                    try
                    {
                        instance = new Midi2Plugin();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Exception message: {ex.Message}");
                        if (ex.InnerException != null)
                        {
                            Debug.WriteLine($"Inner exception message: {ex.InnerException.Message}");
                        }
                        Debug.WriteLine(ex.ToString());
                    }
                }

                return instance;
            }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        private Midi2Plugin()
        {
            Debug.WriteLine("Midi2Plugin contructor called.");
            var installedVersion = Microsoft.Windows.Devices.Midi2.Utilities.RuntimeInformation.MidiRuntimeInformation.GetInstalledVersion();
            Debug.WriteLine($"Installed MIDI runtime version: {installedVersion.Major}.{installedVersion.Minor}, Build: {installedVersion.BuildNumber}, Patch: {installedVersion.Patch}");

            if (midi2Session == null)
            {
                midi2Session = MidiSession.Create("MIDIPluginSession");
            }

            midi2DeviceWatcher = MidiEndpointDeviceWatcher.Create(MidiEndpointDeviceInformationFilters.StandardNativeUniversalMidiPacketFormat);
            midi2DeviceWatcher.Added += UmpDeviceAdded;
            midi2DeviceWatcher.Updated += UmpDeviceUpdated;
            midi2DeviceWatcher.Removed += UmpDeviceRemoved;
            midi2DeviceWatcher.Start();
            Debug.WriteLine("Midi2Plugin initialized and device watcher started.");
        }

        ~Midi2Plugin()
        {
            if (midi2Session != null)
            {
                midi2Session.Dispose();
                midi2Session = null;
            }
            if (midi2DeviceWatcher != null)
            {
                midi2DeviceWatcher.Added -= UmpDeviceAdded;
                midi2DeviceWatcher.Updated -= UmpDeviceUpdated;
                midi2DeviceWatcher.Removed -= UmpDeviceRemoved;
                midi2DeviceWatcher.Stop();
                midi2DeviceWatcher = null;
            }
            Debug.WriteLine("Midi2Plugin finalized and device watcher stopped.");
        }

        private void UmpDeviceAdded(MidiEndpointDeviceWatcher sender, MidiEndpointDeviceInformationAddedEventArgs args)
        {
            Debug.WriteLine($"UmpDeviceAdded device: {args.AddedDevice}, deviceId: {args.AddedDevice.EndpointDeviceId}, instanceId: {args.AddedDevice.DeviceInstanceId}");
            OnMidi2InputDeviceAttached?.Invoke(args.AddedDevice.EndpointDeviceId);
            OnMidi2OutputDeviceAttached?.Invoke(args.AddedDevice.EndpointDeviceId);

            var addedEndpointConnection = midi2Session.CreateEndpointConnection(args.AddedDevice.EndpointDeviceId);
            addedEndpointConnection.MessageReceived += UmpMessageReceived;

            addedEndpointConnection.Open();
            midi2Connections.Add(args.AddedDevice.EndpointDeviceId, addedEndpointConnection);
        }

        private void UmpMessageReceived(IMidiMessageReceivedEventSource sender, MidiMessageReceivedEventArgs midiEvent)
        {
            switch (midiEvent.MessageType)
            {
                case MidiMessageType.UtilityMessage32:
                    Debug.WriteLine($"UtilityMessage32: {midiEvent}, from: {sender}");
                    {
                        var message32 = new MidiMessage32();
                        midiEvent.FillMessage32(message32);
                        OnMidi2UmpReceived?.Invoke(sender.GetEndpointConnectionSource().ConnectedEndpointDeviceId, message32.GetAllWords().ToArray());
                    }
                    break;
                case MidiMessageType.SystemCommon32:
                    Debug.WriteLine($"SystemCommon32: {midiEvent}, from: {sender}");
                    {
                        var message32 = new MidiMessage32();
                        midiEvent.FillMessage32(message32);
                        OnMidi2UmpReceived?.Invoke(sender.GetEndpointConnectionSource().ConnectedEndpointDeviceId, message32.GetAllWords().ToArray());
                    }
                    break;
                case MidiMessageType.Midi1ChannelVoice32:
                    Debug.WriteLine($"Midi1ChannelVoice32: {midiEvent}, from: {sender}");
                    {
                        var message32 = new MidiMessage32();
                        midiEvent.FillMessage32(message32);
                        OnMidi2UmpReceived?.Invoke(sender.GetEndpointConnectionSource().ConnectedEndpointDeviceId, message32.GetAllWords().ToArray());
                    }
                    break;
                case MidiMessageType.DataMessage64:
                    Debug.WriteLine($"DataMessage64: {midiEvent}, from: {sender}");
                    {
                        var message64 = new MidiMessage64();
                        midiEvent.FillMessage64(message64);
                        OnMidi2UmpReceived?.Invoke(sender.GetEndpointConnectionSource().ConnectedEndpointDeviceId, message64.GetAllWords().ToArray());
                    }
                    break;
                case MidiMessageType.Midi2ChannelVoice64:
                    Debug.WriteLine($"Midi2ChannelVoice64: {midiEvent}, from: {sender}");
                    {
                        var message64 = new MidiMessage64();
                        midiEvent.FillMessage64(message64);
                        OnMidi2UmpReceived?.Invoke(sender.GetEndpointConnectionSource().ConnectedEndpointDeviceId, message64.GetAllWords().ToArray());
                    }
                    break;
                case MidiMessageType.DataMessage128:
                    Debug.WriteLine($"DataMessage128: {midiEvent}, from: {sender}");
                    {
                        var message128 = new MidiMessage128();
                        midiEvent.FillMessage128(message128);
                        OnMidi2UmpReceived?.Invoke(sender.GetEndpointConnectionSource().ConnectedEndpointDeviceId, message128.GetAllWords().ToArray());
                    }
                    break;
                case MidiMessageType.FlexData128:
                    Debug.WriteLine($"FlexData128: {midiEvent}, from: {sender}");
                    {
                        var message128 = new MidiMessage128();
                        midiEvent.FillMessage128(message128);
                        OnMidi2UmpReceived?.Invoke(sender.GetEndpointConnectionSource().ConnectedEndpointDeviceId, message128.GetAllWords().ToArray());
                    }
                    break;
                case MidiMessageType.Stream128:
                    Debug.WriteLine($"Stream128: {midiEvent}, from: {sender}");
                    {
                        var message128 = new MidiMessage128();
                        midiEvent.FillMessage128(message128);
                        OnMidi2UmpReceived?.Invoke(sender.GetEndpointConnectionSource().ConnectedEndpointDeviceId, message128.GetAllWords().ToArray());
                    }
                    break;
                default:
                    Debug.WriteLine($"Unknown message type: {midiEvent.MessageType}, message: {midiEvent}, from: {sender}");
                    break;
            }
        }

        /// <summary>
        /// Sends a Universal MIDI Packet (UMP) message to the specified device using its unique identifier.
        /// </summary>
        /// <remarks>If the 'umps' array is null or empty, no message is sent. The method logs a debug
        /// message if the device ID does not correspond to an active connection or if the length of the 'umps' array is
        /// invalid.</remarks>
        /// <param name="deviceId">The unique identifier of the MIDI device to which the UMP message will be sent.</param>
        /// <param name="umps">An array of unsigned integers representing the UMP message data. The array must contain 1, 2, or 4 elements;
        /// otherwise, the message will not be sent.</param>
        public void SendUmpMessage(string deviceId, [ReadOnlyArray] uint[] umps)
        {
            if (umps == null || umps.Length == 0)
            {
                return;
            }

            if (midi2Connections.TryGetValue(deviceId, out var connection))
            {
                var message = new MidiMessageStruct();
                if (umps.Length == 1)
                {
                    message.Word0 = umps[0];
                    connection.SendSingleMessagePacket(MidiMessage32.CreateFromStruct(0, message));
                }
                else if (umps.Length == 2)
                {
                    message.Word0 = umps[0];
                    message.Word1 = umps[1];
                    connection.SendSingleMessagePacket(MidiMessage64.CreateFromStruct(0, message));
                }
                else if (umps.Length == 4)
                {
                    message.Word0 = umps[0];
                    message.Word1 = umps[1];
                    message.Word2 = umps[2];
                    message.Word3 = umps[3];
                    connection.SendSingleMessagePacket(MidiMessage128.CreateFromStruct(0, message));
                }
                else
                {
                    Debug.WriteLine($"SendUmpMessage: Invalid UMP length: {umps.Length}, deviceId: {deviceId}");
                }
            }
            else
            {
                Debug.WriteLine($"SendUmpMessage: No connection found for deviceId: {deviceId}");
            }
        }


        private void UmpDeviceUpdated(MidiEndpointDeviceWatcher sender, MidiEndpointDeviceInformationUpdatedEventArgs args)
        {
            Debug.WriteLine($"UmpDeviceUpdated deviceId: {args.EndpointDeviceId}");
        }

        private void UmpDeviceRemoved(MidiEndpointDeviceWatcher sender, MidiEndpointDeviceInformationRemovedEventArgs args)
        {
            Debug.WriteLine($"UmpDeviceRemoved deviceId: {args.EndpointDeviceId}");
            OnMidi2InputDeviceDetached?.Invoke(args.EndpointDeviceId);
            OnMidi2OutputDeviceDetached?.Invoke(args.EndpointDeviceId);
            if (midi2Connections.TryGetValue(args.EndpointDeviceId, out var connection))
            {
                connection.MessageReceived -= UmpMessageReceived;
                midi2Session.DisconnectEndpointConnection(connection.ConnectionId);
                midi2Connections.Remove(args.EndpointDeviceId);
            }
        }
    }
}