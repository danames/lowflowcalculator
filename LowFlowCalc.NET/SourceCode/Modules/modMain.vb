Module modMain
    Public m_Version As Integer = 1
    Public m_AvePeriod As Integer = 7
    Public m_RetPeriod As Integer = 10
    Public m_ExType As ExtremeTypes = ExtremeTypes.LowFlow
    Public m_ExtremeTypeString As String = "low"
    Public m_DistType As DistTypes = DistTypes.LogPearsonTypeIII
    Public m_DistTypeString As String = "Log Pearson Type III"
    Public m_StartMonth As Integer = 1
    Public m_EndMonth As Integer = 12
    Public m_ShowWelcome As Boolean = True
    Public m_ZerosMethod As Integer = 0 '0 = no method, 1 = remove before averaging, 2 = remove after averaging
    Public m_Prob As Double

    Public Enum ExtremeTypes As Integer
        HighFlow = 0
        LowFlow = 1
    End Enum

    Public Enum DistTypes As Integer
        LogNormal = 0
        LogPearsonTypeIII = 1
    End Enum

    Public Function GetStandardNormalDeviate(ByVal Prob As Double) As Double
        'This function uses an estimation approach to computing the standard normal deviate 
        'and is used to compute the inverse normal probability
        'Taken from Kite page 49
        Dim c0 As Double, c1 As Double, c2 As Double
        Dim d1 As Double, d2 As Double, d3 As Double
        Dim w As Double, Result As Double, P As Double
        c0 = 2.515517
        c1 = 0.802853
        c2 = 0.010328
        d1 = 1.432788
        d2 = 0.189269
        d3 = 0.001308
        If Prob > 0.5 Then
            P = 1 - Prob
        Else
            P = Prob
        End If
        w = Math.Sqrt(Math.Log(1 / P ^ 2))
        Result = w - (c0 + c1 * w + c2 * w ^ 2) / (1 + d1 * w + d2 * w ^ 2 + d3 * w ^ 3)
        If Prob > 0.5 Then
            Return -Result
        Else
            Return Result
        End If
    End Function

    Public Function GetInvNorm(ByVal Prob As Double, ByVal Mean As Double, ByVal StdDev As Double) As Double
        'This function mimics the NormInv function in Excel, returning the size of an event based on the
        'mean and standard deviation of the event probability distribution
        Dim Z As Double
        Z = GetStandardNormalDeviate(Prob)
        Return Mean - Z * StdDev
    End Function

    Public Function GetMean(ByVal InArray As Object) As Double
        'This function returns the mean of an input double array.
        Dim i As Integer
        Dim mn As Double
        mn = 0
        For i = 0 To UBound(InArray)
            mn = mn + InArray(i)
        Next
        mn = mn / (UBound(InArray) + 1)
        mn = Math.Round(mn, 2)
        Return mn
    End Function

    Public Function GetStdDev(ByVal InArray As Object) As Double
        'This function computes the standard deviation of the array data and returns a double
        Dim i As Integer, SumXsq As Double, SumX As Double
        Dim n As Integer, StdDev As Double
        n = UBound(InArray) + 1
        For i = 0 To n - 1
            SumXsq = SumXsq + InArray(i) ^ 2
            SumX = SumX + InArray(i)
        Next
        StdDev = Math.Sqrt((n * SumXsq - SumX ^ 2) / (n * (n - 1)))
        Return StdDev
    End Function


End Module
