'***************************************************************************
'While the underlying libraries are covered by LGPL, this sample is released 
'as public domain.  It is distributed in the hope that it will be useful, but 
'WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY 
'or FITNESS FOR A PARTICULAR PURPOSE.  
'****************************************************************************


Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Collections
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Diagnostics
Imports System.Windows.Forms

Imports DirectShowLib


Namespace SnapShot
    ''' <summary> Summary description for MainForm. </summary>
    Friend Class Capture
        Implements ISampleGrabberCB
        Implements IDisposable

#Region "Member variables"

        ''' <summary> graph builder interface. </summary>
        Private m_FilterGraph As IFilterGraph2 = Nothing

        ' Used to snap picture on Still pin
        Private m_VidControl As IAMVideoControl = Nothing
        Private m_pinStill As IPin = Nothing

        ''' <summary> so we can wait for the async job to finish </summary>
        Private m_PictureReady As ManualResetEvent = Nothing

        Private m_WantOne As Boolean = False

        ''' <summary> Dimensions of the image, calculated once in constructor for perf. </summary>
        Private m_videoWidth As Integer
        Private m_videoHeight As Integer
        Private m_stride As Integer

        ''' <summary> buffer for bitmap data.  Always release by caller</summary>
        Private m_ipBuffer As IntPtr = IntPtr.Zero

#If DEBUG Then
        ' Allow you to "Connect to remote graph" from GraphEdit
        Private m_rot As DsROTEntry = Nothing
#End If
#End Region

#Region "APIs"
        <DllImport("Kernel32.dll", EntryPoint:="RtlMoveMemory")> _
        Private Shared Sub CopyMemory(Destination As IntPtr, Source As IntPtr, <MarshalAs(UnmanagedType.U4)> Length As Integer)
        End Sub
#End Region

        ' Zero based device index and device params and output window
        Public Sub New(iDeviceNum As Integer, iWidth As Integer, iHeight As Integer, iBPP As Short, hControl As Control)
            Dim capDevices As DsDevice()

            ' Get the collection of video devices
            capDevices = DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice)

            If iDeviceNum + 1 > capDevices.Length Then
                Throw New Exception("No video capture devices found at that index!")
            End If

            Try
                ' Set up the capture graph
                SetupGraph(capDevices(iDeviceNum), iWidth, iHeight, iBPP, hControl)

                ' tell the callback to ignore new images
                m_PictureReady = New ManualResetEvent(False)
            Catch
                Dispose()
                Throw
            End Try
        End Sub

        ''' <summary> release everything. </summary>
        Public Sub Dispose()
#If DEBUG Then
            If m_rot IsNot Nothing Then
                m_rot.Dispose()
            End If
#End If
            CloseInterfaces()
            If m_PictureReady IsNot Nothing Then
                m_PictureReady.Close()
            End If
        End Sub
        ' Destructor
        Protected Overrides Sub Finalize()
            Try
                Dispose()
            Finally
                MyBase.Finalize()
            End Try
        End Sub

        ''' <summary>
        ''' Get the image from the Still pin.  The returned image can turned into a bitmap with
        ''' Bitmap b = new Bitmap(cam.Width, cam.Height, cam.Stride, PixelFormat.Format24bppRgb, m_ip);
        ''' If the image is upside down, you can fix it with
        ''' b.RotateFlip(RotateFlipType.RotateNoneFlipY);
        ''' </summary>
        ''' <returns>Returned pointer to be freed by caller with Marshal.FreeCoTaskMem</returns>
        Public Function Click() As IntPtr
            Dim hr As Integer

            ' get ready to wait for new image
            m_PictureReady.Reset()
            m_ipBuffer = Marshal.AllocCoTaskMem(Math.Abs(m_stride) * m_videoHeight)

            Try
                m_WantOne = True

                ' If we are using a still pin, ask for a picture
                If m_VidControl IsNot Nothing Then
                    ' Tell the camera to send an image
                    hr = m_VidControl.SetMode(m_pinStill, VideoControlFlags.Trigger)
                    DsError.ThrowExceptionForHR(hr)
                End If

                ' Start waiting
                If Not m_PictureReady.WaitOne(9000, False) Then
                    Throw New Exception("Timeout waiting to get picture")
                End If
            Catch
                Marshal.FreeCoTaskMem(m_ipBuffer)
                m_ipBuffer = IntPtr.Zero
                Throw
            End Try

            ' Got one
            Return m_ipBuffer
        End Function

        Public ReadOnly Property Width() As Integer
            Get
                Return m_videoWidth
            End Get
        End Property
        Public ReadOnly Property Height() As Integer
            Get
                Return m_videoHeight
            End Get
        End Property
        Public ReadOnly Property Stride() As Integer
            Get
                Return m_stride
            End Get
        End Property


        ''' <summary> build the capture graph for grabber. </summary>
        Private Sub SetupGraph(dev As DsDevice, iWidth As Integer, iHeight As Integer, iBPP As Short, hControl As Control)
            Dim hr As Integer

            Dim sampGrabber As ISampleGrabber = Nothing
            Dim capFilter As IBaseFilter = Nothing
            Dim pCaptureOut As IPin = Nothing
            Dim pSampleIn As IPin = Nothing
            Dim pRenderIn As IPin = Nothing

            ' Get the graphbuilder object
            m_FilterGraph = TryCast(New FilterGraph(), IFilterGraph2)

            Try
#If DEBUG Then
                m_rot = New DsROTEntry(m_FilterGraph)
#End If
                ' add the video input device
                hr = m_FilterGraph.AddSourceFilterForMoniker(dev.Mon, Nothing, dev.Name, capFilter)
                DsError.ThrowExceptionForHR(hr)

                ' Find the still pin
                m_pinStill = DsFindPin.ByCategory(capFilter, PinCategory.Still, 0)

                ' Didn't find one.  Is there a preview pin?
                If m_pinStill Is Nothing Then
                    m_pinStill = DsFindPin.ByCategory(capFilter, PinCategory.Preview, 0)
                End If

                ' Still haven't found one.  Need to put a splitter in so we have
                ' one stream to capture the bitmap from, and one to display.  Ok, we
                ' don't *have* to do it that way, but we are going to anyway.
                If m_pinStill Is Nothing Then
                    Dim pRaw As IPin = Nothing
                    Dim pSmart As IPin = Nothing

                    ' There is no still pin
                    m_VidControl = Nothing

                    ' Add a splitter
                    Dim iSmartTee As IBaseFilter = DirectCast(New SmartTee(), IBaseFilter)

                    Try
                        hr = m_FilterGraph.AddFilter(iSmartTee, "SmartTee")
                        DsError.ThrowExceptionForHR(hr)

                        ' Find the find the capture pin from the video device and the
                        ' input pin for the splitter, and connnect them
                        pRaw = DsFindPin.ByCategory(capFilter, PinCategory.Capture, 0)
                        pSmart = DsFindPin.ByDirection(iSmartTee, PinDirection.Input, 0)

                        hr = m_FilterGraph.Connect(pRaw, pSmart)
                        DsError.ThrowExceptionForHR(hr)

                        ' Now set the capture and still pins (from the splitter)
                        m_pinStill = DsFindPin.ByName(iSmartTee, "Preview")
                        pCaptureOut = DsFindPin.ByName(iSmartTee, "Capture")

                        ' If any of the default config items are set, perform the config
                        ' on the actual video device (rather than the splitter)
                        If iHeight + iWidth + iBPP > 0 Then
                            SetConfigParms(pRaw, iWidth, iHeight, iBPP)
                        End If
                    Finally
                        If pRaw IsNot Nothing Then
                            Marshal.ReleaseComObject(pRaw)
                        End If
                        If Not pRaw.Equals(pSmart) Then
                            Marshal.ReleaseComObject(pSmart)
                        End If
                        If Not pRaw.Equals(iSmartTee) Then
                            Marshal.ReleaseComObject(iSmartTee)
                        End If
                    End Try
                Else
                    ' Get a control pointer (used in Click())
                    m_VidControl = TryCast(capFilter, IAMVideoControl)

                    pCaptureOut = DsFindPin.ByCategory(capFilter, PinCategory.Capture, 0)

                    ' If any of the default config items are set
                    If iHeight + iWidth + iBPP > 0 Then
                        SetConfigParms(m_pinStill, iWidth, iHeight, iBPP)
                    End If
                End If

                ' Get the SampleGrabber interface
                sampGrabber = TryCast(New SampleGrabber(), ISampleGrabber)

                ' Configure the sample grabber
                Dim baseGrabFlt As IBaseFilter = TryCast(sampGrabber, IBaseFilter)
                ConfigureSampleGrabber(sampGrabber)
                pSampleIn = DsFindPin.ByDirection(baseGrabFlt, PinDirection.Input, 0)

                ' Get the default video renderer
                Dim pRenderer As IBaseFilter = TryCast(New VideoRendererDefault(), IBaseFilter)
                hr = m_FilterGraph.AddFilter(pRenderer, "Renderer")
                DsError.ThrowExceptionForHR(hr)

                pRenderIn = DsFindPin.ByDirection(pRenderer, PinDirection.Input, 0)

                ' Add the sample grabber to the graph
                hr = m_FilterGraph.AddFilter(baseGrabFlt, "Ds.NET Grabber")
                DsError.ThrowExceptionForHR(hr)

                If m_VidControl Is Nothing Then
                    ' Connect the Still pin to the sample grabber
                    hr = m_FilterGraph.Connect(m_pinStill, pSampleIn)
                    DsError.ThrowExceptionForHR(hr)

                    ' Connect the capture pin to the renderer
                    hr = m_FilterGraph.Connect(pCaptureOut, pRenderIn)
                    DsError.ThrowExceptionForHR(hr)
                Else
                    ' Connect the capture pin to the renderer
                    hr = m_FilterGraph.Connect(pCaptureOut, pRenderIn)
                    DsError.ThrowExceptionForHR(hr)

                    ' Connect the Still pin to the sample grabber
                    hr = m_FilterGraph.Connect(m_pinStill, pSampleIn)
                    DsError.ThrowExceptionForHR(hr)
                End If

                ' Learn the video properties
                SaveSizeInfo(sampGrabber)
                ConfigVideoWindow(hControl)

                ' Start the graph
                Dim mediaCtrl As IMediaControl = TryCast(m_FilterGraph, IMediaControl)
                hr = mediaCtrl.Run()
                DsError.ThrowExceptionForHR(hr)
            Finally
                If sampGrabber IsNot Nothing Then
                    Marshal.ReleaseComObject(sampGrabber)
                    sampGrabber = Nothing
                End If
                If pCaptureOut IsNot Nothing Then
                    Marshal.ReleaseComObject(pCaptureOut)
                    pCaptureOut = Nothing
                End If
                If pRenderIn IsNot Nothing Then
                    Marshal.ReleaseComObject(pRenderIn)
                    pRenderIn = Nothing
                End If
                If pSampleIn IsNot Nothing Then
                    Marshal.ReleaseComObject(pSampleIn)
                    pSampleIn = Nothing
                End If
            End Try
        End Sub

        Private Sub SaveSizeInfo(sampGrabber As ISampleGrabber)
            Dim hr As Integer

            ' Get the media type from the SampleGrabber
            Dim media As New AMMediaType()

            hr = sampGrabber.GetConnectedMediaType(media)
            DsError.ThrowExceptionForHR(hr)

            If (media.formatType <> FormatType.VideoInfo) OrElse (media.formatPtr = IntPtr.Zero) Then
                Throw New NotSupportedException("Unknown Grabber Media Format")
            End If

            ' Grab the size info
            Dim videoInfoHeader As VideoInfoHeader = DirectCast(Marshal.PtrToStructure(media.formatPtr, GetType(VideoInfoHeader)), VideoInfoHeader)
            m_videoWidth = videoInfoHeader.BmiHeader.Width
            m_videoHeight = videoInfoHeader.BmiHeader.Height
            m_stride = m_videoWidth * (videoInfoHeader.BmiHeader.BitCount / 8)

            DsUtils.FreeAMMediaType(media)
            media = Nothing
        End Sub

        ' Set the video window within the control specified by hControl
        Private Sub ConfigVideoWindow(hControl As Control)
            Dim hr As Integer

            Dim ivw As IVideoWindow = TryCast(m_FilterGraph, IVideoWindow)

            ' Set the parent
            hr = ivw.put_Owner(hControl.Handle)
            DsError.ThrowExceptionForHR(hr)

            ' Turn off captions, etc
            hr = ivw.put_WindowStyle(WindowStyle.Child Or WindowStyle.ClipChildren Or WindowStyle.ClipSiblings)
            DsError.ThrowExceptionForHR(hr)

            ' Yes, make it visible
            hr = ivw.put_Visible(OABool.[True])
            DsError.ThrowExceptionForHR(hr)

            ' Move to upper left corner
            Dim rc As Rectangle = hControl.ClientRectangle
            hr = ivw.SetWindowPosition(0, 0, rc.Right, rc.Bottom)
            DsError.ThrowExceptionForHR(hr)
        End Sub

        Private Sub ConfigureSampleGrabber(sampGrabber As ISampleGrabber)
            Dim hr As Integer
            Dim media As New AMMediaType()

            ' Set the media type to Video/RBG24
            media.majorType = MediaType.Video
            media.subType = MediaSubType.RGB24
            media.formatType = FormatType.VideoInfo
            hr = sampGrabber.SetMediaType(media)
            DsError.ThrowExceptionForHR(hr)

            DsUtils.FreeAMMediaType(media)
            media = Nothing

            ' Configure the samplegrabber
            hr = sampGrabber.SetCallback(Me, 1)
            DsError.ThrowExceptionForHR(hr)
        End Sub

        ' Set the Framerate, and video size
        Private Sub SetConfigParms(pStill As IPin, iWidth As Integer, iHeight As Integer, iBPP As Short)
            Dim hr As Integer
            Dim media As AMMediaType
            Dim v As VideoInfoHeader

            Dim videoStreamConfig As IAMStreamConfig = TryCast(pStill, IAMStreamConfig)

            ' Get the existing format block
            hr = videoStreamConfig.GetFormat(media)
            DsError.ThrowExceptionForHR(hr)

            Try
                ' copy out the videoinfoheader
                v = New VideoInfoHeader()
                Marshal.PtrToStructure(media.formatPtr, v)

                ' if overriding the width, set the width
                If iWidth > 0 Then
                    v.BmiHeader.Width = iWidth
                End If

                ' if overriding the Height, set the Height
                If iHeight > 0 Then
                    v.BmiHeader.Height = iHeight
                End If

                ' if overriding the bits per pixel
                If iBPP > 0 Then
                    v.BmiHeader.BitCount = iBPP
                End If

                ' Copy the media structure back
                Marshal.StructureToPtr(v, media.formatPtr, False)

                ' Set the new format
                hr = videoStreamConfig.SetFormat(media)
                DsError.ThrowExceptionForHR(hr)
            Finally
                DsUtils.FreeAMMediaType(media)
                media = Nothing
            End Try
        End Sub

        ''' <summary> Shut down capture </summary>
        Private Sub CloseInterfaces()
            Dim hr As Integer

            Try
                If m_FilterGraph IsNot Nothing Then
                    Dim mediaCtrl As IMediaControl = TryCast(m_FilterGraph, IMediaControl)

                    ' Stop the graph
                    hr = mediaCtrl.[Stop]()
                End If
            Catch ex As Exception
                Debug.WriteLine(ex)
            End Try

            If m_FilterGraph IsNot Nothing Then
                Marshal.ReleaseComObject(m_FilterGraph)
                m_FilterGraph = Nothing
            End If

            If m_VidControl IsNot Nothing Then
                Marshal.ReleaseComObject(m_VidControl)
                m_VidControl = Nothing
            End If

            If m_pinStill IsNot Nothing Then
                Marshal.ReleaseComObject(m_pinStill)
                m_pinStill = Nothing
            End If
        End Sub

        ''' <summary> sample callback, NOT USED. </summary>
        Private Function ISampleGrabberCB_SampleCB(SampleTime As Double, pSample As IMediaSample) As Integer Implements ISampleGrabberCB.SampleCB
            Marshal.ReleaseComObject(pSample)
            Return 0
        End Function

        ''' <summary> buffer callback, COULD BE FROM FOREIGN THREAD. </summary>
        Private Function ISampleGrabberCB_BufferCB(SampleTime As Double, pBuffer As IntPtr, BufferLen As Integer) As Integer Implements ISampleGrabberCB.BufferCB
            ' Note that we depend on only being called once per call to Click.  Otherwise
            ' a second call can overwrite the previous image.
            Debug.Assert(BufferLen = Math.Abs(m_stride) * m_videoHeight, "Incorrect buffer length")

            If m_WantOne Then
                m_WantOne = False
                Debug.Assert(m_ipBuffer <> IntPtr.Zero, "Unitialized buffer")

                ' Save the buffer
                CopyMemory(m_ipBuffer, pBuffer, BufferLen)

                ' Picture is ready.
                m_PictureReady.[Set]()
            End If

            Return 0
        End Function

        Public Sub Dispose1() Implements IDisposable.Dispose

        End Sub
    End Class
End Namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================
