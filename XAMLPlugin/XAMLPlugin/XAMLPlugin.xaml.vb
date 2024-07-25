Imports System.ComponentModel
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Media.Animation
Imports vMixInterop
Imports System.Linq
Imports Microsoft.Win32
Imports System.Windows.Markup
Imports System.IO
Imports System.Xml
Imports Microsoft.VisualBasic.CompilerServices
Imports System.Linq.Expressions
Imports System.Globalization
Imports System.Runtime.CompilerServices

Namespace XAMLPlugin
    Partial Public Class Title
        Inherits UserControl
        Implements vMixInterop.vMixWPFUserControl

        Private activeStoryBoard As Storyboard
        Private currentDuration As TimeSpan = TimeSpan.Zero
        Private _privateResources As ResourceDictionary
        Private currentPosition As TimeSpan = TimeSpan.Zero
        Private _lastPlay As DateTime = DateTime.Now
        Private isPlay As Boolean = False
        Private isPaused As Boolean = False
        Private _name As String = "No_Name"
        Private _version As String = "0.7.25"


        Public Sub New()
            InitializeComponent()
            Dim dllPath As String = Assembly.GetExecutingAssembly().Location
            Dim pathNoExt As String = Path.Combine(Path.GetDirectoryName(dllPath), Path.GetFileNameWithoutExtension(dllPath))
            Me._name = Path.GetFileNameWithoutExtension(dllPath)
            Me.Log("Ver #" & Me._version)
            Me.Log("Initializing...")
            Dim xamlPath As String = pathNoExt & ".xaml"
            If File.Exists(xamlPath) Then
                Dim loadedUserControl As UserControl = LoadXamlFile(xamlPath)

                loadedUserControl.DataContext = Me.DataContext
                Dim dataContext As Object = loadedUserControl.DataContext
                Dim rootElement As FrameworkElement = CType(loadedUserControl.Content, FrameworkElement)

                Me.Content = Nothing

                Dim Container As New Grid() With {
                            .Name = "ext_XAMLContainer"
                        }

                Me.Content = Container

                If FindElementByName(Me, "txt_Transition") Is Nothing Then
                    Dim txtTransition As TextBlock = New TextBlock With {
                        .Name = "txt_Transition",
                        .Text = "",
                        .Foreground = New SolidColorBrush(Colors.Transparent),
                        .Margin = New Thickness(1920, 0, 0, 0),
                        .Visibility = Visibility.Collapsed
                    }
                    Me.Content.Children.Add(txtTransition)
                End If

                Dim LoadedContent As Object = loadedUserControl.Content
                loadedUserControl.Content = Nothing

                Container.Children.Add(CType(LoadedContent, UIElement))

                ' ========================== RESOURCE HANDLE =========================
                Me._privateResources = loadedUserControl.Resources
                ' ====================================================================

                Me.activeStoryBoard = GetStoryboardByNumber(0)
            Else
                Dim textInfo As TextBlock = New TextBlock With {
                    .Name = "txt_info",
                    .Text = "XAML File: " & Path.GetFileNameWithoutExtension(dllPath) & ".xaml not exist...",
                    .Foreground = New SolidColorBrush(Colors.White),
                    .Margin = New Thickness(500, 500, 0, 0)
                }
                Me.Content.Children.Add(textInfo)
            End If

        End Sub



        Sub Load(width As Integer, height As Integer) Implements vMixInterop.vMixWPFUserControl.Load
            Me.Log("Load: " & width.ToString() & ":" & height.ToString())
            RebuildBindings(Me, 0)
        End Sub

        Public Sub Play() Implements vMixInterop.vMixWPFUserControl.Play
            Me.Log("Play()")
            Dim transitionTextBlock As TextBlock = CType(FindElementByName(Me, "txt_Transition"), TextBlock)
            If transitionTextBlock IsNot Nothing Then
                Dim newActiveStoryboqardName = transitionTextBlock.Text
                Dim newStoryboard As Storyboard
                If newActiveStoryboqardName <> "" Then
                    newStoryboard = GetStoryboard(newActiveStoryboqardName)
                Else
                    newStoryboard = GetStoryboardByNumber(0)
                End If
                If newStoryboard IsNot Nothing Then
                    If Me.activeStoryBoard.GetHashCode() <> newStoryboard.GetHashCode() Then
                        Me.currentPosition = TimeSpan.Zero
                        Me.isPaused = False
                        Me.isPlay = False
                        Me.activeStoryBoard = newStoryboard
                    End If

                Else
                    Me.Log("No storyboard to play")
                    Me.activeStoryBoard = GetStoryboardByNumber(0)
                End If
                startActiveStoryboard()
            End If
        End Sub


        Public Sub Pause() Implements vMixInterop.vMixWPFUserControl.Pause
            Me.Log("Pause()")
            If Me.activeStoryBoard IsNot Nothing Then
                Me.activeStoryBoard.Pause(Me)
                If isPlay Then
                    Me.isPaused = True
                Else
                    Me.isPaused = False
                End If
                Me.isPlay = False
            End If
        End Sub


        Public Sub SetPosition(position As TimeSpan) Implements vMixInterop.vMixWPFUserControl.SetPosition
            Me.Log("SetPosition( " & position.ToString() & " )")
            If Me.activeStoryBoard IsNot Nothing Then
                Me.activeStoryBoard.SeekAlignedToLastTick(Me, position, TimeSeekOrigin.BeginTime)
            End If
            Me.currentPosition = position
        End Sub

        Public Function GetPosition() As TimeSpan Implements vMixInterop.vMixWPFUserControl.GetPosition
            If Me.isPlay Then
                Dim currentNow As DateTime = DateTime.Now
                Dim fromLastPlay As TimeSpan = currentNow - Me._lastPlay
                Me._lastPlay = currentNow
                Me.currentPosition = Me.currentPosition + fromLastPlay
                If Me.currentPosition > currentDuration Then
                    Me.Pause()
                End If
            End If
            Return Me.currentPosition
        End Function


        Public Function GetDuration() As TimeSpan Implements vMixInterop.vMixWPFUserControl.GetDuration
            Return Me.currentDuration
        End Function

        ' Token: 0x0600001E RID: 30
        Public Sub Close() Implements vMixInterop.vMixWPFUserControl.Close
            Me.Log("Close()")
        End Sub

        ' Token: 0x0600001F RID: 31
        Public Sub ShowProperties() Implements vMixInterop.vMixWPFUserControl.ShowProperties
            Me.Log("ShowProperties()")
            MessageBox.Show("XAML Plugin ver." & Me._version & Environment.NewLine & "Written by Nikolay Semenov." & Environment.NewLine & "Contact: semenov_nick@mail.ru")
        End Sub

        Private Function GetStoryboardByNumber(number As Integer) As Storyboard
            If Me._privateResources IsNot Nothing Then
                Dim index As Integer = 0
                For Each value As Object In Me._privateResources.Keys
                    Dim key As String = Conversions.ToString(value)
                    Dim objectValue As Object = RuntimeHelpers.GetObjectValue(Me._privateResources(key))
                    If TypeOf objectValue Is Storyboard Then
                        If index = number Then
                            Return objectValue
                        Else
                            index += 1
                        End If
                    End If
                Next
            End If
            Return Nothing
        End Function

        Private Function LoadXamlFile(filePath As String) As UserControl
            Try
                Me.Log("Reading file from: " & filePath)
                Dim parserContext As ParserContext = New ParserContext() With {
                    .BaseUri = New Uri(filePath)
                    }
                Dim xmlDocument As XmlDocument = New XmlDocument()
                xmlDocument.Load(filePath)
                '======================== IGNORING x:Class ATRRIBUTE ============================
                For i As Integer = xmlDocument.DocumentElement.Attributes.Count - 1 To 0 Step -1
                    Dim xmlAttribute As XmlAttribute = xmlDocument.DocumentElement.Attributes(i)
                    If Operators.CompareString(xmlAttribute.Name, "x:Class", False) = 0 Then
                        xmlDocument.DocumentElement.Attributes.RemoveAt(i)
                    End If
                Next
                '=================================================================================
                Using memoryStream As MemoryStream = New MemoryStream()
                    xmlDocument.Save(memoryStream)
                    memoryStream.Position = 0L

                    Dim LoadedObject As Object = XamlReader.Load(memoryStream, parserContext)

                    If TypeOf LoadedObject Is UserControl Then
                        Return CType(LoadedObject, UserControl)
                    End If
                End Using
            Catch ex As Exception
                MessageBox.Show("Error loading XAML file: " & ex.Message)
            End Try
            Return Nothing
        End Function


        Private Function GetStoryboard(Key As String) As Storyboard
            Dim result As Storyboard
            result = TryCast(_privateResources(Key), Storyboard)
            If result IsNot Nothing Then
                result.FillBehavior = FillBehavior.HoldEnd
            End If
            Return result
        End Function

        Private Sub startActiveStoryboard()
            If Me.activeStoryBoard IsNot Nothing Then
                'AddHandler Me.activeStoryBoard.Completed, AddressOf onStoryboardCompleted
                If Me.isPaused Then
                    Me.activeStoryBoard.Resume(Me)
                Else
                    Me.activeStoryBoard.Begin(Me, True)
                    Me.activeStoryBoard.SeekAlignedToLastTick(Me, Me.currentPosition, TimeSeekOrigin.BeginTime)
                End If
                Me.isPlay = True
                Me._lastPlay = DateTime.Now
                Me.isPaused = False
                Me.currentDuration = Me.GetStoryboardDuration(Me.activeStoryBoard)
            Else
                Me.Pause()
            End If
        End Sub

        Private Function GetStoryboardDuration(storyboard As Storyboard) As TimeSpan
            Dim maxDuration As TimeSpan = TimeSpan.Zero

            For Each timeline As Timeline In storyboard.Children
                Dim endTime As TimeSpan = GetTimelineEndTime(timeline)
                If endTime > maxDuration Then
                    maxDuration = endTime
                End If
            Next

            Return maxDuration
        End Function

        Private Function GetTimelineEndTime(timeline As Timeline) As TimeSpan
            Dim endTime As TimeSpan = TimeSpan.Zero
            ' Console.WriteLine("timeline is {0}", timeline.GetType().Name)
            If TypeOf timeline Is AnimationTimeline Then
                ' Console.WriteLine("AnimationTimeline")
                Dim animation As AnimationTimeline = DirectCast(timeline, AnimationTimeline)
                If animation.Duration.HasTimeSpan Then
                    endTime = animation.BeginTime.GetValueOrDefault() + animation.Duration.TimeSpan
                Else
                    If HasKeyFrames(timeline) Then
                        ' Console.WriteLine("{0} contains keyframes.", timeline.GetType().Name)
                        Dim keyFrames = GetKeyFrames(timeline)
                        ' Find the last key frame time
                        Dim lastKeyFrameTime As TimeSpan = TimeSpan.Zero
                        For Each keyFrame As IKeyFrame In keyFrames
                            If keyFrame.KeyTime.TimeSpan > lastKeyFrameTime Then
                                lastKeyFrameTime = keyFrame.KeyTime.TimeSpan
                            End If
                        Next
                        ' Calculate end time of this animation
                        endTime = animation.BeginTime.GetValueOrDefault() + lastKeyFrameTime
                    Else
                        ' Console.WriteLine("{0} NOT contains keyframes.", timeline.GetType().Name)
                    End If
                End If
            ElseIf TypeOf timeline Is ParallelTimeline Then
                ' Console.WriteLine("ParallelTimeline")
                Dim parallelTimeline As ParallelTimeline = DirectCast(timeline, ParallelTimeline)
                For Each childTimeline As Timeline In parallelTimeline.Children
                    Dim childEndTime As TimeSpan = GetTimelineEndTime(childTimeline)
                    If childEndTime > endTime Then
                        endTime = childEndTime
                    End If
                Next
            ElseIf TypeOf timeline Is Storyboard Then
                ' Console.WriteLine("Storyboard")
                Dim childStoryboard As Storyboard = DirectCast(timeline, Storyboard)
                Dim childEndTime As TimeSpan = GetStoryboardDuration(childStoryboard)
                endTime = childEndTime

            End If

            Return endTime
        End Function
        Private Function HasKeyFrames(animation As Timeline) As Boolean
            'Console.WriteLine("checking Keyframes for {0}", animation.GetType().Name)
            ' Use reflection to get the KeyFrames property
            Dim keyFramesProperty As PropertyInfo = animation.GetType().GetProperty("KeyFrames")
            If keyFramesProperty IsNot Nothing Then
                ' Get the value of the KeyFrames property
                Dim keyFrames As Object = keyFramesProperty.GetValue(animation, Nothing)
                ' Check if the KeyFrames collection contains any keyframes
                Return CType(keyFrames, System.Collections.ICollection).Count > 0
            End If
            Return False
        End Function
        Private Function GetKeyFrames(animation As AnimationTimeline) As IKeyFrame()
            ' Use reflection to get the KeyFrames property
            Dim keyFramesProperty As PropertyInfo = animation.GetType().GetProperty("KeyFrames")
            If keyFramesProperty IsNot Nothing Then
                ' Get the value of the KeyFrames property
                Dim keyFrames As Object = keyFramesProperty.GetValue(animation, Nothing)
                ' Convert the keyframes collection to an array of IKeyFrame
                Return (CType(keyFrames, IEnumerable)).Cast(Of IKeyFrame).ToArray()
            End If
            ' Return an empty IKeyFrame array if there are no keyframes
            Return New IKeyFrame() {}
        End Function

        ' ==========================================================================================================

        Private Function FindElementByName(root As DependencyObject, name As String) As FrameworkElement
            If root Is Nothing Then Return Nothing

            ' Check if the current root is a FrameworkElement and matches the name
            If TypeOf root Is FrameworkElement Then
                Dim element = CType(root, FrameworkElement)
                If element.Name = name Then
                    Return element
                End If
            End If

            ' Check the logical children
            For Each child As Object In LogicalTreeHelper.GetChildren(root)
                If TypeOf child Is DependencyObject Then
                    Dim foundElement = FindElementByName(CType(child, DependencyObject), name)
                    If foundElement IsNot Nothing Then
                        Return foundElement
                    End If
                End If
            Next

            ' Not found
            Return Nothing
        End Function

        ' ===================================================================================================================
        Public Sub RebuildBindings(node As DependencyObject, indent As Integer)
            If node Is Nothing Then
                Return
            End If

            Dim indentString As String = New String(" "c, indent)

            Dim nodeName As String = ""
            Dim hash As String = ""
            If TypeOf node Is FrameworkElement Then
                nodeName = CType(node, FrameworkElement).Name
                hash = CType(node, FrameworkElement).GetHashCode().ToString()
                If nodeName <> "" Then
                    Me.RegisterName(nodeName, CType(node, FrameworkElement))
                End If
            End If

            ' Recurse for each child element
            For Each child As Object In LogicalTreeHelper.GetChildren(node)
                ' Console.WriteLine("childrenCount " & child.GetType().ToString())
                If TypeOf child Is DependencyObject Then
                    RebuildBindings(child, indent + 2)
                End If

            Next
        End Sub
        ' ================================================================================================
        Private Sub onStoryboardCompleted(sender As Object, e As EventArgs)
            Dim storyboard As Storyboard = CType(sender, Storyboard)
            storyboard.Pause(Me)
        End Sub
        Private Sub Log(message As String)
            Console.WriteLine("XAMLPlugin (" & Me._name & "): " & message)
        End Sub
    End Class
End Namespace