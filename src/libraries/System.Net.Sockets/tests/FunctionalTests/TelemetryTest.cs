﻿// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Tracing;
using System.Linq;
using System.Net.Test.Common;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.DotNet.RemoteExecutor;
using Microsoft.DotNet.XUnitExtensions;
using Xunit;
using Xunit.Abstractions;

namespace System.Net.Sockets.Tests
{
    public class TelemetryTest
    {
        private const string ActivitySourceName = "Experimental.System.Net.Sockets";
        private const string ActivityName = ActivitySourceName + ".Connect";

        private static readonly Lazy<Task<bool>> s_remoteServerIsReachable = new Lazy<Task<bool>>(() => Task.Run(async () =>
        {
            try
            {
                using var socket = new Socket(SocketType.Stream, ProtocolType.Tcp);
                EndPoint endPoint = await GetRemoteEndPointAsync(useDnsEndPointString: "True", port: 443);

                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
                await socket.ConnectAsync(endPoint, cts.Token);

                return true;
            }
            catch
            {
                return false;
            }
        }));

        public readonly ITestOutputHelper _output;

        public TelemetryTest(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        [ActiveIssue("https://github.com/dotnet/runtime/issues/107981", TestPlatforms.Wasi)]
        public static void EventSource_ExistsWithCorrectId()
        {
            Type esType = typeof(Socket).Assembly.GetType("System.Net.Sockets.SocketsTelemetry", throwOnError: true, ignoreCase: false);
            Assert.NotNull(esType);

            Assert.Equal("System.Net.Sockets", EventSource.GetName(esType));
            Assert.Equal(Guid.Parse("d5b2e7d4-b6ec-50ae-7cde-af89427ad21f"), EventSource.GetGuid(esType));

            Assert.NotEmpty(EventSource.GenerateManifest(esType, esType.Assembly.Location));
        }

        public static IEnumerable<object[]> SocketMethods_MemberData()
        {
            if (!OperatingSystem.IsWasi()) yield return new[] { "Sync" };
            yield return new[] { "Task" };
            if (!OperatingSystem.IsWasi()) yield return new[] { "Apm" };
            yield return new[] { "Eap" };
        }

        public static IEnumerable<object[]> SocketMethods_Matrix_MemberData()
        {
            return from connectMethod in SocketMethods_MemberData()
                   from acceptMethod in SocketMethods_MemberData()
                   select new[] { connectMethod[0], acceptMethod[0] };
        }

        public static IEnumerable<object[]> SocketMethods_WithBools_MemberData()
        {
            return from connectMethod in SocketMethods_MemberData()
                   from boolValue in new[] { true, false }
                   select new[] { connectMethod[0], boolValue };
        }

        private static async Task<EndPoint> GetRemoteEndPointAsync(string useDnsEndPointString, int port)
        {
            const string Address = "microsoft.com";

            if (bool.Parse(useDnsEndPointString))
            {
                return new DnsEndPoint(Address, port);
            }
            else
            {
                IPAddress ip = (await Dns.GetHostAddressesAsync(Address))[0];
                return new IPEndPoint(ip, port);
            }
        }

        // RemoteExecutor only supports simple argument types such as strings
        // That's why we use this helper method instead of returning SocketHelperBases from MemberDatas directly
        private static SocketHelperBase GetHelperBase(string socketMethod)
        {
            return socketMethod switch
            {
                "Sync" => new SocketHelperArraySync(),
                "Task" => new SocketHelperTask(),
                "Apm" => new SocketHelperApm(),
                "Eap" => new SocketHelperEap(),
                _ => throw new ArgumentException(socketMethod)
            };
        }

        [ConditionalTheory(typeof(RemoteExecutor), nameof(RemoteExecutor.IsSupported))]
        [MemberData(nameof(SocketMethods_WithBools_MemberData))]
        public async Task Connect_Success_ActivityRecorded(string connectMethod, bool ipv6)
        {
            if (ipv6 && !Socket.OSSupportsIPv6) return;

            await RemoteExecutor.Invoke(static async (connectMethod, ipv6Str) =>
            {
                bool ipv6 = bool.Parse(ipv6Str);
                using Socket server = new Socket(ipv6 ? AddressFamily.InterNetworkV6 : AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                server.BindToAnonymousPort(ipv6 ? IPAddress.IPv6Loopback : IPAddress.Loopback);
                server.Listen();

                Activity parent = new Activity("parent").Start();

                using Socket client = new Socket(ipv6 ? AddressFamily.InterNetworkV6 : AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

                using ActivityRecorder recorder = new ActivityRecorder(ActivitySourceName, ActivityName)
                {
                    ExpectedParent = parent
                };

                Task connectTask = GetHelperBase(connectMethod).ConnectAsync(client, server.LocalEndPoint);
                await server.AcceptAsync();
                await connectTask;

                recorder.VerifyActivityRecorded(1);
                Activity activity = recorder.LastFinishedActivity;
                VerifyTcpConnectActivity(activity, (IPEndPoint)server.LocalEndPoint, ipv6);

                Assert.Same(parent, Activity.Current);
                parent.Stop();
            }, connectMethod, ipv6.ToString()).DisposeAsync();
        }

        static void VerifyTcpConnectActivity(Activity activity, IPEndPoint remoteEndPoint, bool ipv6)
        {
            string address = remoteEndPoint.Address.ToString();
            int port = remoteEndPoint.Port;
            Assert.Equal(ActivityKind.Internal, activity.Kind);
            Assert.Equal(ActivityName, activity.OperationName);
            Assert.Equal($"socket connect {address}:{port}", activity.DisplayName);
            ActivityAssert.HasTag(activity, "network.peer.address", address);
            ActivityAssert.HasTag(activity, "network.peer.port", port);
            ActivityAssert.HasTag(activity, "network.type", ipv6 ? "ipv6" : "ipv4");
            ActivityAssert.HasTag(activity, "network.transport", "tcp");
        }

        [OuterLoop("Connection failure takes long on Windows.")]
        [ConditionalTheory(typeof(RemoteExecutor), nameof(RemoteExecutor.IsSupported))]
        [MemberData(nameof(SocketMethods_WithBools_MemberData))]
        public async Task Connect_Failure_ActivityRecorded(string connectMethod, bool ipv6)
        {
            await RemoteExecutor.Invoke(static async (connectMethod, ipv6Str) =>
            {
                bool ipv6 = bool.Parse(ipv6Str);
                using Socket notListening = new Socket(ipv6 ? AddressFamily.InterNetworkV6 : AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                notListening.BindToAnonymousPort(ipv6 ? IPAddress.IPv6Loopback : IPAddress.Loopback);

                Activity parent = new Activity("parent").Start();

                using Socket client = new Socket(ipv6 ? AddressFamily.InterNetworkV6 : AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

                using ActivityRecorder recorder = new ActivityRecorder(ActivitySourceName, ActivityName)
                {
                    ExpectedParent = parent
                };

                SocketException ex = await Assert.ThrowsAsync<SocketException>(() => GetHelperBase(connectMethod)
                    .ConnectAsync(client, notListening.LocalEndPoint));

                recorder.VerifyActivityRecorded(1);
                Activity activity = recorder.LastFinishedActivity;
                VerifyTcpConnectActivity(activity, (IPEndPoint)notListening.LocalEndPoint, ipv6);
                string expectedErrorType = ActivityAssert.CamelToSnake(ex.SocketErrorCode.ToString());
                ActivityAssert.HasTag(activity, "error.type", expectedErrorType);
                Assert.Equal(ActivityStatusCode.Error, activity.Status);

                Assert.Same(parent, Activity.Current);
                parent.Stop();
            }, connectMethod, ipv6.ToString()).DisposeAsync();
        }

        [ConditionalTheory(typeof(RemoteExecutor), nameof(RemoteExecutor.IsSupported))]
        [SkipOnPlatform(TestPlatforms.LinuxBionic, "SElinux blocks UNIX sockets in our CI environment")]
        [MemberData(nameof(SocketMethods_MemberData))]
        public async Task Socket_UDS_Success_ActivityRecorded(string connectMethod)
        {
            if (!Socket.OSSupportsUnixDomainSockets)
            {
                return;
            }

            await RemoteExecutor.Invoke(static async connectMethod =>
            {
                Socket server = null;
                UnixDomainSocketEndPoint endPoint = null;

                //Path selection is contingent on a successful Bind().
                //If it fails, the next iteration will try another path.
                RetryHelper.Execute(() =>
                {
                    server = new Socket(AddressFamily.Unix, SocketType.Stream, ProtocolType.Unspecified);
                    endPoint = new UnixDomainSocketEndPoint(UnixDomainSocketTest.GetRandomNonExistingFilePath());
                    server.Bind(endPoint);
                    server.Listen();
                }, retryWhen: e => e is SocketException);

                using Socket client = new Socket(AddressFamily.Unix, SocketType.Stream, ProtocolType.Unspecified);

                using ActivityRecorder recorder = new ActivityRecorder(ActivitySourceName, ActivityName);

                Task connectTask = GetHelperBase(connectMethod).ConnectAsync(client, endPoint);
                await server.AcceptAsync();
                await connectTask;

                recorder.VerifyActivityRecorded(1);
                Activity activity = recorder.LastFinishedActivity;
                Assert.Equal(ActivityKind.Internal, activity.Kind);
                Assert.Equal(ActivityName, activity.OperationName);
                Assert.Equal($"socket connect {endPoint}", activity.DisplayName);
                ActivityAssert.HasTag(activity, "network.peer.address", endPoint.ToString());
                ActivityAssert.HasTag(activity, "network.transport", "unix");

            }, connectMethod).DisposeAsync();
        }

        [OuterLoop]
        [ConditionalTheory(typeof(RemoteExecutor), nameof(RemoteExecutor.IsSupported))]
        [MemberData(nameof(SocketMethods_Matrix_MemberData))]
        public async Task EventSource_SocketConnectsLoopback_LogsConnectAcceptStartStop(string connectMethod, string acceptMethod)
        {
            await RemoteExecutor.Invoke(async (connectMethod, acceptMethod) =>
            {
                using var listener = new TestEventListener("System.Net.Sockets", EventLevel.Verbose, 0.1);
                listener.AddActivityTracking();

                var events = new ConcurrentQueue<(EventWrittenEventArgs Event, Guid ActivityId)>();
                await listener.RunWithCallbackAsync(e =>
                {
                    events.Enqueue((e, e.ActivityId));

                    if (e.EventName == "ConnectStart")
                    {
                        // Make sure we observe a non-zero current-outgoing-connect-attempts counter
                        WaitForEventCountersAsync(events).GetAwaiter().GetResult();
                    }
                },
                async () =>
                {
                    using var server = new Socket(SocketType.Stream, ProtocolType.Tcp);
                    server.Bind(new IPEndPoint(IPAddress.Loopback, 0));
                    server.Listen();

                    using var client = new Socket(SocketType.Stream, ProtocolType.Tcp);

                    Task connectTask = GetHelperBase(connectMethod).ConnectAsync(client, server.LocalEndPoint);
                    await WaitForEventAsync(events, "ConnectStart");
                    await WaitForEventCountersAsync(events); // Wait for current-outgoing-connect-attempts = 1

                    await GetHelperBase(acceptMethod).AcceptAsync(server);

                    await WaitForEventAsync(events, "AcceptStop");
                    await WaitForEventAsync(events, "ConnectStop");
                    await connectTask;

                    await WaitForEventCountersAsync(events);
                });

                VerifyEvents(events, connect: true, expectedCount: 1);
                VerifyEvents(events, connect: false, expectedCount: 1);
                VerifyEventCounters(events, connectCount: 1, hasCurrentConnectCounter: true);
            }, connectMethod, acceptMethod).DisposeAsync();
        }

        [OuterLoop]
        [ConditionalTheory(typeof(RemoteExecutor), nameof(RemoteExecutor.IsSupported))]
        [MemberData(nameof(SocketMethods_WithBools_MemberData))]
        public async Task EventSource_SocketConnectsRemote_LogsConnectStartStop(string connectMethod, bool useDnsEndPoint)
        {
            if (!await s_remoteServerIsReachable.Value)
            {
                throw new SkipTestException("The remote server is not reachable");
            }

            await RemoteExecutor.Invoke(async (connectMethod, useDnsEndPointString) =>
            {
                using var listener = new TestEventListener("System.Net.Sockets", EventLevel.Verbose, 0.1);
                listener.AddActivityTracking();

                var events = new ConcurrentQueue<(EventWrittenEventArgs Event, Guid ActivityId)>();
                await listener.RunWithCallbackAsync(e => events.Enqueue((e, e.ActivityId)), async () =>
                {
                    using var client = new Socket(SocketType.Stream, ProtocolType.Tcp);

                    SocketHelperBase socketHelper = GetHelperBase(connectMethod);

                    EndPoint endPoint = await GetRemoteEndPointAsync(useDnsEndPointString, port: 443);
                    await socketHelper.ConnectAsync(client, endPoint);

                    await WaitForEventAsync(events, "ConnectStop");

                    await WaitForEventCountersAsync(events);
                });

                VerifyEvents(events, connect: true, expectedCount: 1);
                VerifyEventCounters(events, connectCount: 1, connectOnly: true);
            }, connectMethod, useDnsEndPoint.ToString()).DisposeAsync();
        }

        [OuterLoop]
        [ConditionalTheory(typeof(RemoteExecutor), nameof(RemoteExecutor.IsSupported))]
        [SkipOnPlatform(TestPlatforms.OSX | TestPlatforms.FreeBSD, "Same as Connect.ConnectGetsCanceledByDispose")]
        [MemberData(nameof(SocketMethods_WithBools_MemberData))]
        public async Task EventSource_SocketConnectFailure_LogsConnectFailed(string connectMethod, bool useDnsEndPoint)
        {
            // Skip test on Linux kernels that may have a regression that was fixed in 6.6.
            // See TcpReceiveSendGetsCanceledByDispose test for additional information.
            if (connectMethod == "Sync" && PlatformDetection.IsLinux && Environment.OSVersion.Version < new Version(6, 6))
            {
                return;
            }

            await RemoteExecutor.Invoke(async (connectMethod, useDnsEndPointString) =>
            {
                EndPoint endPoint = await GetRemoteEndPointAsync(useDnsEndPointString, port: 12345);

                using var listener = new TestEventListener("System.Net.Sockets", EventLevel.Verbose, 0.1);
                listener.AddActivityTracking();

                var events = new ConcurrentQueue<(EventWrittenEventArgs Event, Guid ActivityId)>();
                await listener.RunWithCallbackAsync(e => events.Enqueue((e, e.ActivityId)), async () =>
                {
                    using var client = new Socket(SocketType.Stream, ProtocolType.Tcp);

                    SocketHelperBase socketHelper = GetHelperBase(connectMethod);

                    Exception ex = await Assert.ThrowsAnyAsync<Exception>(async () =>
                    {
                        Task connectTask = socketHelper.ConnectAsync(client, endPoint);
                        await WaitForEventAsync(events, "ConnectStart");
                        Task disposeTask = Task.Run(() => client.Dispose());
                        await new[] { connectTask, disposeTask }.WhenAllOrAnyFailed();
                    });

                    if (ex is SocketException se)
                    {
                        Assert.NotEqual(SocketError.TimedOut, se.SocketErrorCode);
                    }

                    await WaitForEventAsync(events, "ConnectStop");

                    await WaitForEventCountersAsync(events);
                });

                // For DNS endpoints, we may see multiple Start/Failure/Stop events
                int? expectedCount = bool.Parse(useDnsEndPointString) ? null : 1;
                VerifyEvents(events, connect: true, expectedCount, shouldHaveFailures: true);
                VerifyEventCounters(events, connectCount: 0);
            }, connectMethod, useDnsEndPoint.ToString()).DisposeAsync();
        }

        [OuterLoop]
        [ConditionalTheory(typeof(RemoteExecutor), nameof(RemoteExecutor.IsSupported))]
        [MemberData(nameof(SocketMethods_MemberData))]
        public async Task EventSource_SocketAcceptFailure_LogsAcceptFailed(string acceptMethod)
        {
            await RemoteExecutor.Invoke(async acceptMethod =>
            {
                using var listener = new TestEventListener("System.Net.Sockets", EventLevel.Verbose, 0.1);
                listener.AddActivityTracking();

                var events = new ConcurrentQueue<(EventWrittenEventArgs Event, Guid ActivityId)>();
                await listener.RunWithCallbackAsync(e => events.Enqueue((e, e.ActivityId)), async () =>
                {
                    using var server = new Socket(SocketType.Stream, ProtocolType.Tcp);
                    server.Bind(new IPEndPoint(IPAddress.Loopback, 0));
                    server.Listen();

                    await Assert.ThrowsAnyAsync<Exception>(async () =>
                    {
                        Task acceptTask = GetHelperBase(acceptMethod).AcceptAsync(server);
                        await WaitForEventAsync(events, "AcceptStart");
                        Task disposeTask = Task.Run(() => server.Dispose());
                        await new[] { acceptTask, disposeTask }.WhenAllOrAnyFailed();
                    });

                    await WaitForEventAsync(events, "AcceptStop");

                    await WaitForEventCountersAsync(events);
                });

                VerifyEvents(events, connect: false, expectedCount: 1, shouldHaveFailures: true);
                VerifyEventCounters(events, connectCount: 0);
            }, acceptMethod).DisposeAsync();
        }

        [OuterLoop]
        [ConditionalTheory(typeof(RemoteExecutor), nameof(RemoteExecutor.IsSupported))]
        [InlineData("Task", true)]
        [InlineData("Task", false)]
        [InlineData("Eap", true)]
        [InlineData("Eap", false)]
        public async Task EventSource_ConnectAsyncCanceled_LogsConnectFailed(string connectMethod, bool useDnsEndPoint)
        {
            await RemoteExecutor.Invoke(async (connectMethod, useDnsEndPointString) =>
            {
                EndPoint endPoint = await GetRemoteEndPointAsync(useDnsEndPointString, port: 12345);

                using var listener = new TestEventListener("System.Net.Sockets", EventLevel.Verbose, 0.1);
                listener.AddActivityTracking();

                var events = new ConcurrentQueue<(EventWrittenEventArgs Event, Guid ActivityId)>();
                await listener.RunWithCallbackAsync(e => events.Enqueue((e, e.ActivityId)), async () =>
                {
                    using var client = new Socket(SocketType.Stream, ProtocolType.Tcp);

                    await Assert.ThrowsAnyAsync<Exception>(async () =>
                    {
                        switch (connectMethod)
                        {
                            case "Task":
                                using (var cts = new CancellationTokenSource())
                                {
                                    ValueTask connectTask = client.ConnectAsync(endPoint, cts.Token);
                                    await WaitForEventAsync(events, "ConnectStart");
                                    cts.Cancel();
                                    await connectTask;
                                }
                                break;

                            case "Eap":
                                using (var saea = new SocketAsyncEventArgs())
                                {
                                    var tcs = new TaskCompletionSource();
                                    saea.RemoteEndPoint = endPoint;
                                    saea.Completed += (_, __) =>
                                    {
                                        Assert.NotEqual(SocketError.Success, saea.SocketError);
                                        tcs.SetException(new SocketException((int)saea.SocketError));
                                    };
                                    Assert.True(client.ConnectAsync(saea));
                                    await WaitForEventAsync(events, "ConnectStart");
                                    Socket.CancelConnectAsync(saea);
                                    await tcs.Task;
                                }
                                break;
                        }
                    });

                    await WaitForEventAsync(events, "ConnectStop");

                    await WaitForEventCountersAsync(events);
                });

                // For DNS endpoints, we may see multiple Start/Failure/Stop events
                int? expectedCount = bool.Parse(useDnsEndPointString) ? null : 1;
                VerifyEvents(events, connect: true, expectedCount, shouldHaveFailures: true);
                VerifyEventCounters(events, connectCount: 0);
            }, connectMethod, useDnsEndPoint.ToString()).DisposeAsync();
        }

        [OuterLoop]
        [ConditionalFact(typeof(RemoteExecutor), nameof(RemoteExecutor.IsSupported))]
        public async Task EventSource_EventsRaisedAsExpected()
        {
            await RemoteExecutor.Invoke(async () =>
            {
                using (var listener = new TestEventListener("System.Net.Sockets", EventLevel.Verbose, 0.1))
                {
                    listener.AddActivityTracking();

                    var events = new ConcurrentQueue<(EventWrittenEventArgs Event, Guid ActivityId)>();
                    await listener.RunWithCallbackAsync(e => events.Enqueue((e, e.ActivityId)), async () =>
                    {
                        // Invoke several tests to execute code paths while tracing is enabled

                        await new SendReceive_Sync(null).SendRecv_Stream_TCP(IPAddress.Loopback, false).ConfigureAwait(false);
                        await new SendReceive_Sync(null).SendRecv_Stream_TCP(IPAddress.Loopback, true).ConfigureAwait(false);

                        await new SendReceive_Task(null).SendRecv_Stream_TCP(IPAddress.Loopback, false).ConfigureAwait(false);
                        await new SendReceive_Task(null).SendRecv_Stream_TCP(IPAddress.Loopback, true).ConfigureAwait(false);

                        await new SendReceive_Eap(null).SendRecv_Stream_TCP(IPAddress.Loopback, false).ConfigureAwait(false);
                        await new SendReceive_Eap(null).SendRecv_Stream_TCP(IPAddress.Loopback, true).ConfigureAwait(false);

                        await new SendReceive_Apm(null).SendRecv_Stream_TCP(IPAddress.Loopback, false).ConfigureAwait(false);
                        await new SendReceive_Apm(null).SendRecv_Stream_TCP(IPAddress.Loopback, true).ConfigureAwait(false);

                        await new SendReceiveUdpClient().SendToRecvFromAsync_Datagram_UDP_UdpClient(IPAddress.Loopback, false).ConfigureAwait(false);
                        await new SendReceiveUdpClient().SendToRecvFromAsync_Datagram_UDP_UdpClient(IPAddress.Loopback, false).ConfigureAwait(false);

                        await new NetworkStreamTest().CopyToAsync_AllDataCopied(4096, true).ConfigureAwait(false);
                        await new NetworkStreamTest().Timeout_Roundtrips().ConfigureAwait(false);

                        await WaitForEventCountersAsync(events);
                    });

                    VerifyEvents(events, connect: true, expectedCount: 10);
                    VerifyEventCounters(events, connectCount: 10, shouldHaveTransferredBytes: true, shouldHaveDatagrams: true);
                }
            }).DisposeAsync();
        }

        private static async Task WaitForEventAsync(ConcurrentQueue<(EventWrittenEventArgs Event, Guid ActivityId)> events, string name)
        {
            DateTime startTime = DateTime.UtcNow;
            while (!events.Any(e => e.Event.EventName == name))
            {
                if (DateTime.UtcNow.Subtract(startTime) > TimeSpan.FromSeconds(30))
                    throw new TimeoutException($"Timed out waiting for {name}");

                await Task.Delay(100);
            }
        }

        private static async Task WaitForEventCountersAsync(ConcurrentQueue<(EventWrittenEventArgs Event, Guid ActivityId)> events)
        {
            DateTime startTime = DateTime.UtcNow;
            int startCount = events.Count;

            while (events.Skip(startCount).Count(e => IsBytesSentEventCounter(e.Event)) < 3)
            {
                if (DateTime.UtcNow.Subtract(startTime) > TimeSpan.FromSeconds(30))
                    throw new TimeoutException($"Timed out waiting for EventCounters");

                await Task.Delay(100);
            }

            static bool IsBytesSentEventCounter(EventWrittenEventArgs e)
            {
                if (e.EventName != "EventCounters")
                    return false;

                var dictionary = (IDictionary<string, object>)e.Payload.Single();

                return (string)dictionary["Name"] == "bytes-sent";
            }
        }

        private static void VerifyEvents(ConcurrentQueue<(EventWrittenEventArgs Event, Guid ActivityId)> events, bool connect, int? expectedCount, bool shouldHaveFailures = false)
        {
            bool start = false;
            Guid startGuid = Guid.Empty;
            bool seenFailures = false;
            bool seenFailureAfterStart = false;
            int numberOfStops = 0;

            foreach ((EventWrittenEventArgs Event, Guid ActivityId) in events)
            {
                Assert.False(Event.EventId == 0, $"Received an error event from EventSource: {Event.Message}");

                if (Event.EventName.Contains("Connect") != connect)
                {
                    continue;
                }

                switch (Event.EventName)
                {
                    case "ConnectStart":
                    case "AcceptStart":
                        Assert.False(start, "Start without a Stop");
                        Assert.NotEqual(Guid.Empty, ActivityId);
                        startGuid = ActivityId;
                        seenFailureAfterStart = false;
                        start = true;

                        string startAddress = Assert.IsType<string>(Assert.Single(Event.Payload));
                        Assert.Matches(@"^InterNetwork.*?:\d\d:{(?:\d{1,3},?)+}$", startAddress);
                        break;

                    case "ConnectStop":
                    case "AcceptStop":
                        Assert.True(start, "Stop without a Start");
                        Assert.Equal(startGuid, ActivityId);
                        startGuid = Guid.Empty;
                        numberOfStops++;
                        start = false;

                        Assert.Empty(Event.Payload);
                        break;

                    case "ConnectFailed":
                    case "AcceptFailed":
                        Assert.True(start, "Failed should come between Start and Stop");
                        Assert.False(seenFailureAfterStart, "Start may only have one Failed event");
                        Assert.Equal(startGuid, ActivityId);
                        seenFailureAfterStart = true;
                        seenFailures = true;

                        Assert.Equal(2, Event.Payload.Count);
                        Assert.True(Enum.IsDefined((SocketError)Event.Payload[0]));
                        Assert.IsType<string>(Event.Payload[1]);
                        break;
                }
            }

            Assert.False(start, "Start without a Stop");
            Assert.Equal(shouldHaveFailures, seenFailures);

            if (expectedCount.HasValue)
            {
                Assert.Equal(expectedCount, numberOfStops);
            }
            else
            {
                Assert.NotEqual(0, numberOfStops);
            }
        }

        private static void VerifyEventCounters(ConcurrentQueue<(EventWrittenEventArgs Event, Guid ActivityId)> events, int connectCount, bool hasCurrentConnectCounter = false, bool connectOnly = false, bool shouldHaveTransferredBytes = false, bool shouldHaveDatagrams = false)
        {
            Dictionary<string, double[]> eventCounters = events
                .Where(e => e.Event.EventName == "EventCounters")
                .Select(e => (IDictionary<string, object>)e.Event.Payload.Single())
                .GroupBy(d => (string)d["Name"], d => (double)(d.ContainsKey("Mean") ? d["Mean"] : d["Increment"]))
                .ToDictionary(p => p.Key, p => p.ToArray());

            Assert.True(eventCounters.TryGetValue("outgoing-connections-established", out double[] outgoingConnections));
            Assert.Equal(connectCount, outgoingConnections[^1]);

            Assert.True(eventCounters.TryGetValue("incoming-connections-established", out double[] incomingConnections));
            Assert.Equal(connectOnly ? 0 : connectCount, incomingConnections[^1]);

            Assert.True(eventCounters.TryGetValue("current-outgoing-connect-attempts", out double[] currentOutgoingConnectAttempts));
            if (hasCurrentConnectCounter)
            {
                Assert.Contains(currentOutgoingConnectAttempts, c => c > 0);
            }
            Assert.Equal(0, currentOutgoingConnectAttempts[^1]);

            Assert.True(eventCounters.TryGetValue("bytes-received", out double[] bytesReceived));
            if (shouldHaveTransferredBytes)
            {
                Assert.True(bytesReceived[^1] > 0);
            }

            Assert.True(eventCounters.TryGetValue("bytes-sent", out double[] bytesSent));
            if (shouldHaveTransferredBytes)
            {
                Assert.True(bytesSent[^1] > 0);
            }

            Assert.True(eventCounters.TryGetValue("datagrams-received", out double[] datagramsReceived));
            if (shouldHaveDatagrams)
            {
                Assert.True(datagramsReceived[^1] > 0);
            }

            Assert.True(eventCounters.TryGetValue("datagrams-sent", out double[] datagramsSent));
            if (shouldHaveDatagrams)
            {
                Assert.True(datagramsSent[^1] > 0);
            }
        }
    }
}
