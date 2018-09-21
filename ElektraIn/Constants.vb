Public Class Constants
    Public Shared ReadOnly ElektraConnectionNamespace As String = "EFLanguage.LS.Production"
End Class
Public NotInheritable Class Products
    Public Shared ReadOnly School As String = "SCH" 'LTBooking
    Public Shared ReadOnly LanguageSchools As String = "LS" 'LSBooking
    Public Shared ReadOnly AcademicProgramsAndPathways As String = "APP" 'APPBooking
    Public Shared ReadOnly LanguageSchoolsProfessionals As String = "LSP" 'LSPBooking
    'Public Shared ReadOnly BritinAcademy As String = "BA" 'BABooking
End Class
Public NotInheritable Class Programs
    Public Shared ReadOnly InternationalLanguageCamps As String = "ILC" 'SCH 
    Public Shared ReadOnly InternationalLanguageSchools As String = "ILS" 'LS
    Public Shared ReadOnly PreparationWeeks As String = "PRW" 'APP 
    Public Shared ReadOnly MultipleLanguageYear As String = "MLY" 'APP
    Public Shared ReadOnly AcademicYear As String = "AY" 'APP
    Public Shared ReadOnly InternationalBaccalaureate As String = "IB" 'IA
    Public Shared ReadOnly IBPrep As String = "IBP" 'IA
    Public Shared ReadOnly ALevels As String = "AL" 'BA
    Public Shared ReadOnly InternationalBaccalaureatePreparation As String = "IAP"
    Public Shared ReadOnly InternationalGerneralCertificate = "GCSE"

    'ELEK-4537
    Public Shared ReadOnly InternationalLanguageSchools25Plus = "ILSP"
    Public Shared ReadOnly AcademicYearProfessionals = "AYP"
    Public Shared ReadOnly MultiLanguageYearProfessionals = "MLYP"
    Public Shared ReadOnly UniversityPreparationProfessionals = "UPP"

    Public Shared ReadOnly EFCorporate = "EFC"

End Class
Public NotInheritable Class SystemArticleTypeCodes
    Public Shared ReadOnly Extension As String = "EX"
    Public Shared ReadOnly Termination As String = "TE"
    Public Shared ReadOnly Transfer As String = "TR"
End Class