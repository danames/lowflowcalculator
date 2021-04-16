VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLowFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Data() As Single          'the array of daily observations
Private m_StartDate As Date         'start date for the time series
Private m_EndMonth As Integer       'end of the period to look at
Private m_StartMonth As Integer     'start of the period to look at
Private m_NumPoints As Long         'number of points in the full series
Private m_AvgPeriod As Integer      'averaging period they want.. ie 7 days
Private m_RetPeriod As Integer      'return period of interest.. ie 10 years
Private m_Averages() As Single      'the data averaged by the averaging period
Private m_YearMins() As Single      'the mins of the averages within the start and stop months
Private m_Years() As Integer        'the years associated with the above
Private m_FirstRun As Boolean       'used for redimming the yearmins array
Private m_FullPeriods As Boolean    'indicates whether only full periods are used
'

Public Sub LoadData(TimeSeries As Variant, StartDate As Date)
    'This function assumes that the timeseries is and array of daily values
    'and that the startdate is the start date...
    m_Data = TimeSeries
    m_StartDate = StartDate
    m_NumPoints = UBound(TimeSeries) + 1
End Sub

Public Property Get NumPoints() As Long
    NumPoints = m_NumPoints
End Property

Public Function Compute(AvgPeriod As Integer, RetPeriod As Integer, FullPeriods As Boolean) As Boolean
    If NumPoints = 0 Then
        Compute = False
        Exit Function
    End If
    m_AvgPeriod = AvgPeriod
    m_RetPeriod = RetPeriod
    
    If FullPeriods = True Then
        TrimToPeriod
    End If
    GetAverages
    GetYearMins
    GetEstimate
End Function

Private Sub GetEstimate()
    Dim i As Long
    Dim Sd As Single, Mn As Single, Sm As Single, n As Long
    Sm = 0
    n = UBound(m_YearMins) + 1
    For i = 0 To n - 1
        Sm = Sm + m_YearMins(i)
    Next i
    Mn = Sm / n
    Sm = 0
    For i = 0 To n - 1
        Sm = Sm + (m_YearMins(i) - Mn) ^ 2
    Next i
    Sd = Sqr(Sm / (n - 1))
    
End Sub

Private Sub TrimToPeriod()
    Dim dStep As Date, StartIndex As Long, EndIndex As Long
    Dim NewData() As Single, i As Long
    dStep = m_StartDate - 1
    For i = 1 To m_NumPoints
        dStep = dStep + 1
        If Month(dStep) = m_StartMonth Then
            StartIndex = i
            Exit For
        End If
    Next i
    For i = m_NumPoints To 1 Step -1
        dStep = m_StartDate + i
        If Month(dStep) = m_EndMonth Then
            EndIndex = i
            Exit For
        End If
    Next i
    ReDim NewData(EndIndex - StartIndex + 1)
    For i = 0 To EndIndex - StartIndex + 1
        NewData(i) = m_Data(i + StartIndex - 1)
    Next i
    Erase m_Data
    m_Data = NewData
    m_StartDate = m_StartDate + StartIndex
    m_NumPoints = EndIndex - StartIndex + 2
End Sub

Public Function ComputeBoot(AvgPeriod As Integer, RetPeriod As Integer, NumBoots As Integer) As Boolean
    If NumPoints = 0 Then
        ComputeBoot = False
        Exit Function
    End If
    m_AvgPeriod = AvgPeriod
    m_RetPeriod = RetPeriod
End Function


Private Function GetAverages() As Boolean
    'get the averages for point for the next m_AvgPeriod days for full record
    Dim i As Long, j As Long
    ReDim m_Averages(m_NumPoints - 1 - m_AvgPeriod)
    For i = 0 To m_NumPoints - 1 - m_AvgPeriod
        For j = i To i + m_AvgPeriod - 1
            m_Averages(i) = m_Averages(i) + m_Data(j)
        Next j
        m_Averages(i) = m_Averages(i) / m_AvgPeriod
    Next i
End Function

Private Function GetYearMins()
    Dim dStep As Date, CrossYear As Boolean, nYears As Long
    Dim mStep1 As Integer, mStep2 As Integer, yStep As Integer
    Dim i As Long
    If m_StartMonth > m_EndMonth Then
        CrossYear = True
    Else
        CrossYear = False
    End If
    nYears = 0
    dStep = m_StartDate - 1
    For i = 0 To UBound(m_Averages)
        dStep = dStep + 1
        mStep1 = Month(dStep)
        mStep2 = Month(dStep + m_AvgPeriod)
        yStep = Year(dStep)
        If CrossYear = False Then
            If mStep1 >= m_StartMonth And mStep2 <= m_EndMonth Then
                UpdateMins yStep, m_Averages(i)
            End If
        Else
            If mStep1 >= m_StartMonth Or mStep2 <= m_EndMonth Then
                UpdateMins yStep, m_Averages(i)
            End If
        End If
    Next i
End Function

Private Sub UpdateMins(vYear As Integer, vNewValue As Single)
    Dim Found As Boolean, i As Long
    If m_FirstRun = True Then
        ReDim Preserve m_Years(0)
        m_Years(0) = vYear
        ReDim Preserve m_YearMins(0)
        m_YearMins(0) = vNewValue
        m_FirstRun = False
        Exit Sub
    End If
    Found = False
    For i = 0 To UBound(m_Years)
        If m_Years(i) = vYear Then
            Found = True
            If vNewValue < m_YearMins(i) Then
                m_YearMins(i) = vNewValue
                Exit For
            End If
        End If
    Next i
    If Found = False Then 'new year
        ReDim Preserve m_Years(UBound(m_Years) + 1)
        m_Years(UBound(m_Years)) = vYear
        ReDim Preserve m_YearMins(UBound(m_YearMins) + 1)
        m_YearMins(UBound(m_YearMins)) = vNewValue
    End If
End Sub

Public Sub Clear()
    'erase all data and reset everything
    'dpa 9/25/02
    Erase m_Data
    m_StartDate = 0
    m_EndMonth = 0
    m_StartMonth = 0
    m_NumPoints = 0
    m_AvgPeriod = 0
    m_RetPeriod = 0
    Erase m_Averages
    Erase m_YearMins
    Erase m_Years
    m_FirstRun = True
End Sub

Public Property Get StartMonth() As Variant
    StartMonth = m_StartMonth
End Property

Public Property Let StartMonth(ByVal vNewValue As Variant)
    m_StartMonth = vNewValue
End Property

Public Property Get EndMonth() As Variant
    EndMonth = m_EndMonth
End Property

Public Property Let EndMonth(ByVal vNewValue As Variant)
    m_EndMonth = vNewValue
End Property

Private Function GetPhi(Z As Single) As Single
    Dim i As Integer
    i = Z * 10
    
    Select Case i
    Case -1
    
End Function
