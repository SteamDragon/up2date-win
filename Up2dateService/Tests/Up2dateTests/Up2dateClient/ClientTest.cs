﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Threading;
using System.Threading.Tasks;
using Tests_Shared;
using Up2dateClient;
using Up2dateShared;

namespace Up2dateTests.Up2dateClient
{
    [TestClass]
    public class ClientTest
    {
        private WrapperMock wrapperMock;
        private SettingsManagerMock settingsManagerMock;
        private SetupManagerMock setupManagerMock;
        private LoggerMock loggerMock;
        private string certificate;

        [TestCleanup]
        public void Cleanup()
        {
            wrapperMock.ExitRun();
        }

        [TestMethod]
        public void WhenCreated_ClientStatusIsStopped()
        {
            // arrange
            // act
            Client client = CreateClient();

            // assert
            Assert.IsNotNull(client.State);
            Assert.AreEqual(ClientStatus.Stopped, client.State.Status);
            Assert.IsTrue(string.IsNullOrEmpty(client.State.LastError));
        }

        [DataTestMethod]
        [DataRow(null)]
        [DataRow("")]
        public void GivenNoCertificate_WhenRun_ThenWrapperRunIsNotExecuted_AndStatusIsNoCertificate(string certificate)
        {
            // arrange
            Client client = CreateClient();
            this.certificate = certificate;

            // act
            client.Run();

            // assert
            wrapperMock.VerifyNoOtherCalls();
            Assert.AreEqual(ClientStatus.NoCertificate, client.State.Status);
        }

        [TestMethod]
        public void WhenRun_WrapperClientRunIsCalledWithCorrectArguments()
        {
            // arrange
            Client client = CreateClient();

            // act
            StartClient(client);

            // assert
            wrapperMock.Verify(m => m.RunClient(certificate, settingsManagerMock.Object.ProvisioningUrl, settingsManagerMock.Object.XApigToken,
                wrapperMock.Dispatcher, It.IsNotNull<AuthErrorActionFunc>()));
        }

        [TestMethod]
        public void GivenClientRunning_WhenWrapperClientRunExited_ThenStatusIsReconnectingWithoutMessage_AndDispatcherIsDeleted()
        {
            // arrange
            Client client = CreateClient();
            StartClient(client);

            // act
            wrapperMock.ExitRun();
            Thread.Sleep(100);

            // assert
            Assert.AreEqual(ClientStatus.Reconnecting, client.State.Status);
            Assert.IsTrue(string.IsNullOrEmpty(client.State.LastError));
            wrapperMock.Verify(m => m.DeleteDispatcher(wrapperMock.Dispatcher));
        }

        [TestMethod]
        public void GivenClientRunning_WhenWrapperClientThrewException_ThenStatusIsReconnectingWithMessage_AndDispatcherIsDeleted()
        {
            // arrange
            Client client = CreateClient();
            string message = "exception message";
            wrapperMock.Setup(m => m.RunClient(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<IntPtr>(), It.IsAny<AuthErrorActionFunc>()))
                .Throws(new Exception(message));

            // act
            client.Run();

            // assert
            Assert.AreEqual(ClientStatus.Reconnecting, client.State.Status);
            Assert.AreEqual(message, client.State.LastError);
            wrapperMock.Verify(m => m.DeleteDispatcher(wrapperMock.Dispatcher));
        }

        [TestMethod]
        public void WhenRun_ThenStatusIsRunning()
        {
            // arrange
            Client client = CreateClient();

            // act
            StartClient(client);

            // assert
            Assert.AreEqual(ClientStatus.Running, client.State.Status);
            Assert.IsTrue(string.IsNullOrEmpty(client.State.LastError));
        }

        [TestMethod]
        public void GivenClientRunning_WhenAuthErrorActionBringsErrorMessage_ThenStatusIsAuthorizationError()
        {
            // arrange
            Client client = CreateClient();
            string message = "Authorization Error message";
            StartClient(client);

            // act
            wrapperMock.AuthErrorCallback(message);

            // assert
            Assert.AreEqual(ClientStatus.AuthorizationError, client.State.Status);
            Assert.AreEqual(message, client.State.LastError);
        }

        private void StartClient(Client client)
        {
            Task.Run(client.Run);
            Thread.Sleep(100);
        }

        private Client CreateClient()
        {
            wrapperMock = new WrapperMock();
            settingsManagerMock = new SettingsManagerMock();
            setupManagerMock = new SetupManagerMock();
            loggerMock = new LoggerMock();
            certificate = "certificate body";
            Client client = new Client(wrapperMock.Object, settingsManagerMock.Object, () => certificate, setupManagerMock.Object, GetSysInfo, loggerMock.Object);

            return client;
        }

        private SystemInfo GetSysInfo()
        {
            throw new NotImplementedException();
        }
    }
}
