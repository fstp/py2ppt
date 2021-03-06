# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 2.7.8 (default, Jun 30 2014, 16:03:49) [MSC v.1500 32 bit (Intel)]
# From type library 'MSO.DLL'
# On Sun Oct 04 23:14:38 2015
'Microsoft Office 12.0 Object Library'
makepy_version = '0.5.01'
python_version = 0x20708f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}')
MajorVersion = 2
MinorVersion = 4
LibraryFlags = 8
LCID = 0x0

class constants:
	certdetAvailable              =0          # from enum CertificateDetail
	certdetExpirationDate         =3          # from enum CertificateDetail
	certdetIssuer                 =2          # from enum CertificateDetail
	certdetSubject                =1          # from enum CertificateDetail
	certdetThumbprint             =4          # from enum CertificateDetail
	certverresError               =0          # from enum CertificateVerificationResults
	certverresExpired             =5          # from enum CertificateVerificationResults
	certverresInvalid             =4          # from enum CertificateVerificationResults
	certverresRevoked             =6          # from enum CertificateVerificationResults
	certverresUntrusted           =7          # from enum CertificateVerificationResults
	certverresUnverified          =2          # from enum CertificateVerificationResults
	certverresValid               =3          # from enum CertificateVerificationResults
	certverresVerifying           =1          # from enum CertificateVerificationResults
	contverresError               =0          # from enum ContentVerificationResults
	contverresModified            =4          # from enum ContentVerificationResults
	contverresUnverified          =2          # from enum ContentVerificationResults
	contverresValid               =3          # from enum ContentVerificationResults
	contverresVerifying           =1          # from enum ContentVerificationResults
	offPropertyTypeBoolean        =2          # from enum DocProperties
	offPropertyTypeDate           =3          # from enum DocProperties
	offPropertyTypeFloat          =5          # from enum DocProperties
	offPropertyTypeNumber         =1          # from enum DocProperties
	offPropertyTypeString         =4          # from enum DocProperties
	cipherModeCBC                 =1          # from enum EncryptionCipherMode
	cipherModeECB                 =0          # from enum EncryptionCipherMode
	encprovdetAlgorithm           =1          # from enum EncryptionProviderDetail
	encprovdetBlockCipher         =2          # from enum EncryptionProviderDetail
	encprovdetCipherBlockSize     =3          # from enum EncryptionProviderDetail
	encprovdetCipherMode          =4          # from enum EncryptionProviderDetail
	encprovdetUrl                 =0          # from enum EncryptionProviderDetail
	mfHTML                        =2          # from enum MailFormat
	mfPlainText                   =1          # from enum MailFormat
	mfRTF                         =3          # from enum MailFormat
	msoAlertButtonAbortRetryIgnore=2          # from enum MsoAlertButtonType
	msoAlertButtonOK              =0          # from enum MsoAlertButtonType
	msoAlertButtonOKCancel        =1          # from enum MsoAlertButtonType
	msoAlertButtonRetryCancel     =5          # from enum MsoAlertButtonType
	msoAlertButtonYesAllNoCancel  =6          # from enum MsoAlertButtonType
	msoAlertButtonYesNo           =4          # from enum MsoAlertButtonType
	msoAlertButtonYesNoCancel     =3          # from enum MsoAlertButtonType
	msoAlertCancelDefault         =-1         # from enum MsoAlertCancelType
	msoAlertCancelFifth           =4          # from enum MsoAlertCancelType
	msoAlertCancelFirst           =0          # from enum MsoAlertCancelType
	msoAlertCancelFourth          =3          # from enum MsoAlertCancelType
	msoAlertCancelSecond          =1          # from enum MsoAlertCancelType
	msoAlertCancelThird           =2          # from enum MsoAlertCancelType
	msoAlertDefaultFifth          =4          # from enum MsoAlertDefaultType
	msoAlertDefaultFirst          =0          # from enum MsoAlertDefaultType
	msoAlertDefaultFourth         =3          # from enum MsoAlertDefaultType
	msoAlertDefaultSecond         =1          # from enum MsoAlertDefaultType
	msoAlertDefaultThird          =2          # from enum MsoAlertDefaultType
	msoAlertIconCritical          =1          # from enum MsoAlertIconType
	msoAlertIconInfo              =4          # from enum MsoAlertIconType
	msoAlertIconNoIcon            =0          # from enum MsoAlertIconType
	msoAlertIconQuery             =2          # from enum MsoAlertIconType
	msoAlertIconWarning           =3          # from enum MsoAlertIconType
	msoAlignBottoms               =5          # from enum MsoAlignCmd
	msoAlignCenters               =1          # from enum MsoAlignCmd
	msoAlignLefts                 =0          # from enum MsoAlignCmd
	msoAlignMiddles               =4          # from enum MsoAlignCmd
	msoAlignRights                =2          # from enum MsoAlignCmd
	msoAlignTops                  =3          # from enum MsoAlignCmd
	msoAnimationAppear            =32         # from enum MsoAnimationType
	msoAnimationBeginSpeaking     =4          # from enum MsoAnimationType
	msoAnimationCharacterSuccessMajor=6          # from enum MsoAnimationType
	msoAnimationCheckingSomething =103        # from enum MsoAnimationType
	msoAnimationDisappear         =31         # from enum MsoAnimationType
	msoAnimationEmptyTrash        =116        # from enum MsoAnimationType
	msoAnimationGestureDown       =113        # from enum MsoAnimationType
	msoAnimationGestureLeft       =114        # from enum MsoAnimationType
	msoAnimationGestureRight      =19         # from enum MsoAnimationType
	msoAnimationGestureUp         =115        # from enum MsoAnimationType
	msoAnimationGetArtsy          =100        # from enum MsoAnimationType
	msoAnimationGetAttentionMajor =11         # from enum MsoAnimationType
	msoAnimationGetAttentionMinor =12         # from enum MsoAnimationType
	msoAnimationGetTechy          =101        # from enum MsoAnimationType
	msoAnimationGetWizardy        =102        # from enum MsoAnimationType
	msoAnimationGoodbye           =3          # from enum MsoAnimationType
	msoAnimationGreeting          =2          # from enum MsoAnimationType
	msoAnimationIdle              =1          # from enum MsoAnimationType
	msoAnimationListensToComputer =26         # from enum MsoAnimationType
	msoAnimationLookDown          =104        # from enum MsoAnimationType
	msoAnimationLookDownLeft      =105        # from enum MsoAnimationType
	msoAnimationLookDownRight     =106        # from enum MsoAnimationType
	msoAnimationLookLeft          =107        # from enum MsoAnimationType
	msoAnimationLookRight         =108        # from enum MsoAnimationType
	msoAnimationLookUp            =109        # from enum MsoAnimationType
	msoAnimationLookUpLeft        =110        # from enum MsoAnimationType
	msoAnimationLookUpRight       =111        # from enum MsoAnimationType
	msoAnimationPrinting          =18         # from enum MsoAnimationType
	msoAnimationRestPose          =5          # from enum MsoAnimationType
	msoAnimationSaving            =112        # from enum MsoAnimationType
	msoAnimationSearching         =13         # from enum MsoAnimationType
	msoAnimationSendingMail       =25         # from enum MsoAnimationType
	msoAnimationThinking          =24         # from enum MsoAnimationType
	msoAnimationWorkingAtSomething=23         # from enum MsoAnimationType
	msoAnimationWritingNotingSomething=22         # from enum MsoAnimationType
	msoLanguageIDExeMode          =4          # from enum MsoAppLanguageID
	msoLanguageIDHelp             =3          # from enum MsoAppLanguageID
	msoLanguageIDInstall          =1          # from enum MsoAppLanguageID
	msoLanguageIDUI               =2          # from enum MsoAppLanguageID
	msoLanguageIDUIPrevious       =5          # from enum MsoAppLanguageID
	msoArrowheadLengthMedium      =2          # from enum MsoArrowheadLength
	msoArrowheadLengthMixed       =-2         # from enum MsoArrowheadLength
	msoArrowheadLong              =3          # from enum MsoArrowheadLength
	msoArrowheadShort             =1          # from enum MsoArrowheadLength
	msoArrowheadDiamond           =5          # from enum MsoArrowheadStyle
	msoArrowheadNone              =1          # from enum MsoArrowheadStyle
	msoArrowheadOpen              =3          # from enum MsoArrowheadStyle
	msoArrowheadOval              =6          # from enum MsoArrowheadStyle
	msoArrowheadStealth           =4          # from enum MsoArrowheadStyle
	msoArrowheadStyleMixed        =-2         # from enum MsoArrowheadStyle
	msoArrowheadTriangle          =2          # from enum MsoArrowheadStyle
	msoArrowheadNarrow            =1          # from enum MsoArrowheadWidth
	msoArrowheadWide              =3          # from enum MsoArrowheadWidth
	msoArrowheadWidthMedium       =2          # from enum MsoArrowheadWidth
	msoArrowheadWidthMixed        =-2         # from enum MsoArrowheadWidth
	msoShape10pointStar           =149        # from enum MsoAutoShapeType
	msoShape12pointStar           =150        # from enum MsoAutoShapeType
	msoShape16pointStar           =94         # from enum MsoAutoShapeType
	msoShape24pointStar           =95         # from enum MsoAutoShapeType
	msoShape32pointStar           =96         # from enum MsoAutoShapeType
	msoShape4pointStar            =91         # from enum MsoAutoShapeType
	msoShape5pointStar            =92         # from enum MsoAutoShapeType
	msoShape6pointStar            =147        # from enum MsoAutoShapeType
	msoShape7pointStar            =148        # from enum MsoAutoShapeType
	msoShape8pointStar            =93         # from enum MsoAutoShapeType
	msoShapeActionButtonBackorPrevious=129        # from enum MsoAutoShapeType
	msoShapeActionButtonBeginning =131        # from enum MsoAutoShapeType
	msoShapeActionButtonCustom    =125        # from enum MsoAutoShapeType
	msoShapeActionButtonDocument  =134        # from enum MsoAutoShapeType
	msoShapeActionButtonEnd       =132        # from enum MsoAutoShapeType
	msoShapeActionButtonForwardorNext=130        # from enum MsoAutoShapeType
	msoShapeActionButtonHelp      =127        # from enum MsoAutoShapeType
	msoShapeActionButtonHome      =126        # from enum MsoAutoShapeType
	msoShapeActionButtonInformation=128        # from enum MsoAutoShapeType
	msoShapeActionButtonMovie     =136        # from enum MsoAutoShapeType
	msoShapeActionButtonReturn    =133        # from enum MsoAutoShapeType
	msoShapeActionButtonSound     =135        # from enum MsoAutoShapeType
	msoShapeArc                   =25         # from enum MsoAutoShapeType
	msoShapeBalloon               =137        # from enum MsoAutoShapeType
	msoShapeBentArrow             =41         # from enum MsoAutoShapeType
	msoShapeBentUpArrow           =44         # from enum MsoAutoShapeType
	msoShapeBevel                 =15         # from enum MsoAutoShapeType
	msoShapeBlockArc              =20         # from enum MsoAutoShapeType
	msoShapeCan                   =13         # from enum MsoAutoShapeType
	msoShapeChartPlus             =182        # from enum MsoAutoShapeType
	msoShapeChartStar             =181        # from enum MsoAutoShapeType
	msoShapeChartX                =180        # from enum MsoAutoShapeType
	msoShapeChevron               =52         # from enum MsoAutoShapeType
	msoShapeChord                 =161        # from enum MsoAutoShapeType
	msoShapeCircularArrow         =60         # from enum MsoAutoShapeType
	msoShapeCloud                 =179        # from enum MsoAutoShapeType
	msoShapeCloudCallout          =108        # from enum MsoAutoShapeType
	msoShapeCorner                =162        # from enum MsoAutoShapeType
	msoShapeCornerTabs            =169        # from enum MsoAutoShapeType
	msoShapeCross                 =11         # from enum MsoAutoShapeType
	msoShapeCube                  =14         # from enum MsoAutoShapeType
	msoShapeCurvedDownArrow       =48         # from enum MsoAutoShapeType
	msoShapeCurvedDownRibbon      =100        # from enum MsoAutoShapeType
	msoShapeCurvedLeftArrow       =46         # from enum MsoAutoShapeType
	msoShapeCurvedRightArrow      =45         # from enum MsoAutoShapeType
	msoShapeCurvedUpArrow         =47         # from enum MsoAutoShapeType
	msoShapeCurvedUpRibbon        =99         # from enum MsoAutoShapeType
	msoShapeDecagon               =144        # from enum MsoAutoShapeType
	msoShapeDiagonalStripe        =141        # from enum MsoAutoShapeType
	msoShapeDiamond               =4          # from enum MsoAutoShapeType
	msoShapeDodecagon             =146        # from enum MsoAutoShapeType
	msoShapeDonut                 =18         # from enum MsoAutoShapeType
	msoShapeDoubleBrace           =27         # from enum MsoAutoShapeType
	msoShapeDoubleBracket         =26         # from enum MsoAutoShapeType
	msoShapeDoubleWave            =104        # from enum MsoAutoShapeType
	msoShapeDownArrow             =36         # from enum MsoAutoShapeType
	msoShapeDownArrowCallout      =56         # from enum MsoAutoShapeType
	msoShapeDownRibbon            =98         # from enum MsoAutoShapeType
	msoShapeExplosion1            =89         # from enum MsoAutoShapeType
	msoShapeExplosion2            =90         # from enum MsoAutoShapeType
	msoShapeFlowchartAlternateProcess=62         # from enum MsoAutoShapeType
	msoShapeFlowchartCard         =75         # from enum MsoAutoShapeType
	msoShapeFlowchartCollate      =79         # from enum MsoAutoShapeType
	msoShapeFlowchartConnector    =73         # from enum MsoAutoShapeType
	msoShapeFlowchartData         =64         # from enum MsoAutoShapeType
	msoShapeFlowchartDecision     =63         # from enum MsoAutoShapeType
	msoShapeFlowchartDelay        =84         # from enum MsoAutoShapeType
	msoShapeFlowchartDirectAccessStorage=87         # from enum MsoAutoShapeType
	msoShapeFlowchartDisplay      =88         # from enum MsoAutoShapeType
	msoShapeFlowchartDocument     =67         # from enum MsoAutoShapeType
	msoShapeFlowchartExtract      =81         # from enum MsoAutoShapeType
	msoShapeFlowchartInternalStorage=66         # from enum MsoAutoShapeType
	msoShapeFlowchartMagneticDisk =86         # from enum MsoAutoShapeType
	msoShapeFlowchartManualInput  =71         # from enum MsoAutoShapeType
	msoShapeFlowchartManualOperation=72         # from enum MsoAutoShapeType
	msoShapeFlowchartMerge        =82         # from enum MsoAutoShapeType
	msoShapeFlowchartMultidocument=68         # from enum MsoAutoShapeType
	msoShapeFlowchartOfflineStorage=139        # from enum MsoAutoShapeType
	msoShapeFlowchartOffpageConnector=74         # from enum MsoAutoShapeType
	msoShapeFlowchartOr           =78         # from enum MsoAutoShapeType
	msoShapeFlowchartPredefinedProcess=65         # from enum MsoAutoShapeType
	msoShapeFlowchartPreparation  =70         # from enum MsoAutoShapeType
	msoShapeFlowchartProcess      =61         # from enum MsoAutoShapeType
	msoShapeFlowchartPunchedTape  =76         # from enum MsoAutoShapeType
	msoShapeFlowchartSequentialAccessStorage=85         # from enum MsoAutoShapeType
	msoShapeFlowchartSort         =80         # from enum MsoAutoShapeType
	msoShapeFlowchartStoredData   =83         # from enum MsoAutoShapeType
	msoShapeFlowchartSummingJunction=77         # from enum MsoAutoShapeType
	msoShapeFlowchartTerminator   =69         # from enum MsoAutoShapeType
	msoShapeFoldedCorner          =16         # from enum MsoAutoShapeType
	msoShapeFrame                 =158        # from enum MsoAutoShapeType
	msoShapeFunnel                =174        # from enum MsoAutoShapeType
	msoShapeGear6                 =172        # from enum MsoAutoShapeType
	msoShapeGear9                 =173        # from enum MsoAutoShapeType
	msoShapeHalfFrame             =159        # from enum MsoAutoShapeType
	msoShapeHeart                 =21         # from enum MsoAutoShapeType
	msoShapeHeptagon              =145        # from enum MsoAutoShapeType
	msoShapeHexagon               =10         # from enum MsoAutoShapeType
	msoShapeHorizontalScroll      =102        # from enum MsoAutoShapeType
	msoShapeIsoscelesTriangle     =7          # from enum MsoAutoShapeType
	msoShapeLeftArrow             =34         # from enum MsoAutoShapeType
	msoShapeLeftArrowCallout      =54         # from enum MsoAutoShapeType
	msoShapeLeftBrace             =31         # from enum MsoAutoShapeType
	msoShapeLeftBracket           =29         # from enum MsoAutoShapeType
	msoShapeLeftCircularArrow     =176        # from enum MsoAutoShapeType
	msoShapeLeftRightArrow        =37         # from enum MsoAutoShapeType
	msoShapeLeftRightArrowCallout =57         # from enum MsoAutoShapeType
	msoShapeLeftRightCircularArrow=177        # from enum MsoAutoShapeType
	msoShapeLeftRightRibbon       =140        # from enum MsoAutoShapeType
	msoShapeLeftRightUpArrow      =40         # from enum MsoAutoShapeType
	msoShapeLeftUpArrow           =43         # from enum MsoAutoShapeType
	msoShapeLightningBolt         =22         # from enum MsoAutoShapeType
	msoShapeLineCallout1          =109        # from enum MsoAutoShapeType
	msoShapeLineCallout1AccentBar =113        # from enum MsoAutoShapeType
	msoShapeLineCallout1BorderandAccentBar=121        # from enum MsoAutoShapeType
	msoShapeLineCallout1NoBorder  =117        # from enum MsoAutoShapeType
	msoShapeLineCallout2          =110        # from enum MsoAutoShapeType
	msoShapeLineCallout2AccentBar =114        # from enum MsoAutoShapeType
	msoShapeLineCallout2BorderandAccentBar=122        # from enum MsoAutoShapeType
	msoShapeLineCallout2NoBorder  =118        # from enum MsoAutoShapeType
	msoShapeLineCallout3          =111        # from enum MsoAutoShapeType
	msoShapeLineCallout3AccentBar =115        # from enum MsoAutoShapeType
	msoShapeLineCallout3BorderandAccentBar=123        # from enum MsoAutoShapeType
	msoShapeLineCallout3NoBorder  =119        # from enum MsoAutoShapeType
	msoShapeLineCallout4          =112        # from enum MsoAutoShapeType
	msoShapeLineCallout4AccentBar =116        # from enum MsoAutoShapeType
	msoShapeLineCallout4BorderandAccentBar=124        # from enum MsoAutoShapeType
	msoShapeLineCallout4NoBorder  =120        # from enum MsoAutoShapeType
	msoShapeLineInverse           =183        # from enum MsoAutoShapeType
	msoShapeMathDivide            =166        # from enum MsoAutoShapeType
	msoShapeMathEqual             =167        # from enum MsoAutoShapeType
	msoShapeMathMinus             =164        # from enum MsoAutoShapeType
	msoShapeMathMultiply          =165        # from enum MsoAutoShapeType
	msoShapeMathNotEqual          =168        # from enum MsoAutoShapeType
	msoShapeMathPlus              =163        # from enum MsoAutoShapeType
	msoShapeMixed                 =-2         # from enum MsoAutoShapeType
	msoShapeMoon                  =24         # from enum MsoAutoShapeType
	msoShapeNoSymbol              =19         # from enum MsoAutoShapeType
	msoShapeNonIsoscelesTrapezoid =143        # from enum MsoAutoShapeType
	msoShapeNotPrimitive          =138        # from enum MsoAutoShapeType
	msoShapeNotchedRightArrow     =50         # from enum MsoAutoShapeType
	msoShapeOctagon               =6          # from enum MsoAutoShapeType
	msoShapeOval                  =9          # from enum MsoAutoShapeType
	msoShapeOvalCallout           =107        # from enum MsoAutoShapeType
	msoShapeParallelogram         =2          # from enum MsoAutoShapeType
	msoShapePentagon              =51         # from enum MsoAutoShapeType
	msoShapePie                   =142        # from enum MsoAutoShapeType
	msoShapePieWedge              =175        # from enum MsoAutoShapeType
	msoShapePlaque                =28         # from enum MsoAutoShapeType
	msoShapePlaqueTabs            =171        # from enum MsoAutoShapeType
	msoShapeQuadArrow             =39         # from enum MsoAutoShapeType
	msoShapeQuadArrowCallout      =59         # from enum MsoAutoShapeType
	msoShapeRectangle             =1          # from enum MsoAutoShapeType
	msoShapeRectangularCallout    =105        # from enum MsoAutoShapeType
	msoShapeRegularPentagon       =12         # from enum MsoAutoShapeType
	msoShapeRightArrow            =33         # from enum MsoAutoShapeType
	msoShapeRightArrowCallout     =53         # from enum MsoAutoShapeType
	msoShapeRightBrace            =32         # from enum MsoAutoShapeType
	msoShapeRightBracket          =30         # from enum MsoAutoShapeType
	msoShapeRightTriangle         =8          # from enum MsoAutoShapeType
	msoShapeRound1Rectangle       =151        # from enum MsoAutoShapeType
	msoShapeRound2DiagRectangle   =153        # from enum MsoAutoShapeType
	msoShapeRound2SameRectangle   =152        # from enum MsoAutoShapeType
	msoShapeRoundedRectangle      =5          # from enum MsoAutoShapeType
	msoShapeRoundedRectangularCallout=106        # from enum MsoAutoShapeType
	msoShapeSmileyFace            =17         # from enum MsoAutoShapeType
	msoShapeSnip1Rectangle        =155        # from enum MsoAutoShapeType
	msoShapeSnip2DiagRectangle    =157        # from enum MsoAutoShapeType
	msoShapeSnip2SameRectangle    =156        # from enum MsoAutoShapeType
	msoShapeSnipRoundRectangle    =154        # from enum MsoAutoShapeType
	msoShapeSquareTabs            =170        # from enum MsoAutoShapeType
	msoShapeStripedRightArrow     =49         # from enum MsoAutoShapeType
	msoShapeSun                   =23         # from enum MsoAutoShapeType
	msoShapeSwooshArrow           =178        # from enum MsoAutoShapeType
	msoShapeTear                  =160        # from enum MsoAutoShapeType
	msoShapeTrapezoid             =3          # from enum MsoAutoShapeType
	msoShapeUTurnArrow            =42         # from enum MsoAutoShapeType
	msoShapeUpArrow               =35         # from enum MsoAutoShapeType
	msoShapeUpArrowCallout        =55         # from enum MsoAutoShapeType
	msoShapeUpDownArrow           =38         # from enum MsoAutoShapeType
	msoShapeUpDownArrowCallout    =58         # from enum MsoAutoShapeType
	msoShapeUpRibbon              =97         # from enum MsoAutoShapeType
	msoShapeVerticalScroll        =101        # from enum MsoAutoShapeType
	msoShapeWave                  =103        # from enum MsoAutoShapeType
	msoAutoSizeMixed              =-2         # from enum MsoAutoSize
	msoAutoSizeNone               =0          # from enum MsoAutoSize
	msoAutoSizeShapeToFitText     =1          # from enum MsoAutoSize
	msoAutoSizeTextToFitShape     =2          # from enum MsoAutoSize
	msoAutomationSecurityByUI     =2          # from enum MsoAutomationSecurity
	msoAutomationSecurityForceDisable=3          # from enum MsoAutomationSecurity
	msoAutomationSecurityLow      =1          # from enum MsoAutomationSecurity
	msoBackgroundStyleMixed       =-2         # from enum MsoBackgroundStyleIndex
	msoBackgroundStyleNotAPreset  =0          # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset1     =1          # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset10    =10         # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset11    =11         # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset12    =12         # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset2     =2          # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset3     =3          # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset4     =4          # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset5     =5          # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset6     =6          # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset7     =7          # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset8     =8          # from enum MsoBackgroundStyleIndex
	msoBackgroundStylePreset9     =9          # from enum MsoBackgroundStyleIndex
	msoBalloonButtonAbort         =-8         # from enum MsoBalloonButtonType
	msoBalloonButtonBack          =-5         # from enum MsoBalloonButtonType
	msoBalloonButtonCancel        =-2         # from enum MsoBalloonButtonType
	msoBalloonButtonClose         =-12        # from enum MsoBalloonButtonType
	msoBalloonButtonIgnore        =-9         # from enum MsoBalloonButtonType
	msoBalloonButtonNext          =-6         # from enum MsoBalloonButtonType
	msoBalloonButtonNo            =-4         # from enum MsoBalloonButtonType
	msoBalloonButtonNull          =0          # from enum MsoBalloonButtonType
	msoBalloonButtonOK            =-1         # from enum MsoBalloonButtonType
	msoBalloonButtonOptions       =-14        # from enum MsoBalloonButtonType
	msoBalloonButtonRetry         =-7         # from enum MsoBalloonButtonType
	msoBalloonButtonSearch        =-10        # from enum MsoBalloonButtonType
	msoBalloonButtonSnooze        =-11        # from enum MsoBalloonButtonType
	msoBalloonButtonTips          =-13        # from enum MsoBalloonButtonType
	msoBalloonButtonYes           =-3         # from enum MsoBalloonButtonType
	msoBalloonButtonYesToAll      =-15        # from enum MsoBalloonButtonType
	msoBalloonErrorBadCharacter   =8          # from enum MsoBalloonErrorType
	msoBalloonErrorBadPictureRef  =4          # from enum MsoBalloonErrorType
	msoBalloonErrorBadReference   =5          # from enum MsoBalloonErrorType
	msoBalloonErrorButtonModeless =7          # from enum MsoBalloonErrorType
	msoBalloonErrorButtonlessModal=6          # from enum MsoBalloonErrorType
	msoBalloonErrorCOMFailure     =9          # from enum MsoBalloonErrorType
	msoBalloonErrorCharNotTopmostForModal=10         # from enum MsoBalloonErrorType
	msoBalloonErrorNone           =0          # from enum MsoBalloonErrorType
	msoBalloonErrorOther          =1          # from enum MsoBalloonErrorType
	msoBalloonErrorOutOfMemory    =3          # from enum MsoBalloonErrorType
	msoBalloonErrorTooBig         =2          # from enum MsoBalloonErrorType
	msoBalloonErrorTooManyControls=11         # from enum MsoBalloonErrorType
	msoBalloonTypeBullets         =1          # from enum MsoBalloonType
	msoBalloonTypeButtons         =0          # from enum MsoBalloonType
	msoBalloonTypeNumbers         =2          # from enum MsoBalloonType
	msoBarBottom                  =3          # from enum MsoBarPosition
	msoBarFloating                =4          # from enum MsoBarPosition
	msoBarLeft                    =0          # from enum MsoBarPosition
	msoBarMenuBar                 =6          # from enum MsoBarPosition
	msoBarPopup                   =5          # from enum MsoBarPosition
	msoBarRight                   =2          # from enum MsoBarPosition
	msoBarTop                     =1          # from enum MsoBarPosition
	msoBarNoChangeDock            =16         # from enum MsoBarProtection
	msoBarNoChangeVisible         =8          # from enum MsoBarProtection
	msoBarNoCustomize             =1          # from enum MsoBarProtection
	msoBarNoHorizontalDock        =64         # from enum MsoBarProtection
	msoBarNoMove                  =4          # from enum MsoBarProtection
	msoBarNoProtection            =0          # from enum MsoBarProtection
	msoBarNoResize                =2          # from enum MsoBarProtection
	msoBarNoVerticalDock          =32         # from enum MsoBarProtection
	msoBarRowFirst                =0          # from enum MsoBarRow
	msoBarRowLast                 =-1         # from enum MsoBarRow
	msoBarTypeMenuBar             =1          # from enum MsoBarType
	msoBarTypeNormal              =0          # from enum MsoBarType
	msoBarTypePopup               =2          # from enum MsoBarType
	msoBaselineAlignAuto          =5          # from enum MsoBaselineAlignment
	msoBaselineAlignBaseline      =1          # from enum MsoBaselineAlignment
	msoBaselineAlignCenter        =3          # from enum MsoBaselineAlignment
	msoBaselineAlignFarEast50     =4          # from enum MsoBaselineAlignment
	msoBaselineAlignMixed         =-2         # from enum MsoBaselineAlignment
	msoBaselineAlignTop           =2          # from enum MsoBaselineAlignment
	msoBevelAngle                 =6          # from enum MsoBevelType
	msoBevelArtDeco               =13         # from enum MsoBevelType
	msoBevelCircle                =3          # from enum MsoBevelType
	msoBevelConvex                =8          # from enum MsoBevelType
	msoBevelCoolSlant             =9          # from enum MsoBevelType
	msoBevelCross                 =5          # from enum MsoBevelType
	msoBevelDivot                 =10         # from enum MsoBevelType
	msoBevelHardEdge              =12         # from enum MsoBevelType
	msoBevelNone                  =1          # from enum MsoBevelType
	msoBevelRelaxedInset          =2          # from enum MsoBevelType
	msoBevelRiblet                =11         # from enum MsoBevelType
	msoBevelSlope                 =4          # from enum MsoBevelType
	msoBevelSoftRound             =7          # from enum MsoBevelType
	msoBevelTypeMixed             =-2         # from enum MsoBevelType
	msoBlackWhiteAutomatic        =1          # from enum MsoBlackWhiteMode
	msoBlackWhiteBlack            =8          # from enum MsoBlackWhiteMode
	msoBlackWhiteBlackTextAndLine =6          # from enum MsoBlackWhiteMode
	msoBlackWhiteDontShow         =10         # from enum MsoBlackWhiteMode
	msoBlackWhiteGrayOutline      =5          # from enum MsoBlackWhiteMode
	msoBlackWhiteGrayScale        =2          # from enum MsoBlackWhiteMode
	msoBlackWhiteHighContrast     =7          # from enum MsoBlackWhiteMode
	msoBlackWhiteInverseGrayScale =4          # from enum MsoBlackWhiteMode
	msoBlackWhiteLightGrayScale   =3          # from enum MsoBlackWhiteMode
	msoBlackWhiteMixed            =-2         # from enum MsoBlackWhiteMode
	msoBlackWhiteWhite            =9          # from enum MsoBlackWhiteMode
	msoBlogMultipleCategories     =2          # from enum MsoBlogCategorySupport
	msoBlogNoCategories           =0          # from enum MsoBlogCategorySupport
	msoBlogOneCategory            =1          # from enum MsoBlogCategorySupport
	msoblogImageTypeGIF           =2          # from enum MsoBlogImageType
	msoblogImageTypeJPEG          =1          # from enum MsoBlogImageType
	msoblogImageTypePNG           =3          # from enum MsoBlogImageType
	msoBulletMixed                =-2         # from enum MsoBulletType
	msoBulletNone                 =0          # from enum MsoBulletType
	msoBulletNumbered             =2          # from enum MsoBulletType
	msoBulletPicture              =3          # from enum MsoBulletType
	msoBulletUnnumbered           =1          # from enum MsoBulletType
	msoButtonSetAbortRetryIgnore  =10         # from enum MsoButtonSetType
	msoButtonSetBackClose         =6          # from enum MsoButtonSetType
	msoButtonSetBackNextClose     =8          # from enum MsoButtonSetType
	msoButtonSetBackNextSnooze    =12         # from enum MsoButtonSetType
	msoButtonSetCancel            =2          # from enum MsoButtonSetType
	msoButtonSetNextClose         =7          # from enum MsoButtonSetType
	msoButtonSetNone              =0          # from enum MsoButtonSetType
	msoButtonSetOK                =1          # from enum MsoButtonSetType
	msoButtonSetOkCancel          =3          # from enum MsoButtonSetType
	msoButtonSetRetryCancel       =9          # from enum MsoButtonSetType
	msoButtonSetSearchClose       =11         # from enum MsoButtonSetType
	msoButtonSetTipsOptionsClose  =13         # from enum MsoButtonSetType
	msoButtonSetYesAllNoCancel    =14         # from enum MsoButtonSetType
	msoButtonSetYesNo             =4          # from enum MsoButtonSetType
	msoButtonSetYesNoCancel       =5          # from enum MsoButtonSetType
	msoButtonDown                 =-1         # from enum MsoButtonState
	msoButtonMixed                =2          # from enum MsoButtonState
	msoButtonUp                   =0          # from enum MsoButtonState
	msoButtonAutomatic            =0          # from enum MsoButtonStyle
	msoButtonCaption              =2          # from enum MsoButtonStyle
	msoButtonIcon                 =1          # from enum MsoButtonStyle
	msoButtonIconAndCaption       =3          # from enum MsoButtonStyle
	msoButtonIconAndCaptionBelow  =11         # from enum MsoButtonStyle
	msoButtonIconAndWrapCaption   =7          # from enum MsoButtonStyle
	msoButtonIconAndWrapCaptionBelow=15         # from enum MsoButtonStyle
	msoButtonWrapCaption          =14         # from enum MsoButtonStyle
	msoButtonTextBelow            =8          # from enum MsoButtonStyleHidden
	msoButtonWrapText             =4          # from enum MsoButtonStyleHidden
	msoCTPDockPositionBottom      =3          # from enum MsoCTPDockPosition
	msoCTPDockPositionFloating    =4          # from enum MsoCTPDockPosition
	msoCTPDockPositionLeft        =0          # from enum MsoCTPDockPosition
	msoCTPDockPositionRight       =2          # from enum MsoCTPDockPosition
	msoCTPDockPositionTop         =1          # from enum MsoCTPDockPosition
	msoCTPDockPositionRestrictNoChange=1          # from enum MsoCTPDockPositionRestrict
	msoCTPDockPositionRestrictNoHorizontal=2          # from enum MsoCTPDockPositionRestrict
	msoCTPDockPositionRestrictNoVertical=3          # from enum MsoCTPDockPositionRestrict
	msoCTPDockPositionRestrictNone=0          # from enum MsoCTPDockPositionRestrict
	msoCalloutAngle30             =2          # from enum MsoCalloutAngleType
	msoCalloutAngle45             =3          # from enum MsoCalloutAngleType
	msoCalloutAngle60             =4          # from enum MsoCalloutAngleType
	msoCalloutAngle90             =5          # from enum MsoCalloutAngleType
	msoCalloutAngleAutomatic      =1          # from enum MsoCalloutAngleType
	msoCalloutAngleMixed          =-2         # from enum MsoCalloutAngleType
	msoCalloutDropBottom          =4          # from enum MsoCalloutDropType
	msoCalloutDropCenter          =3          # from enum MsoCalloutDropType
	msoCalloutDropCustom          =1          # from enum MsoCalloutDropType
	msoCalloutDropMixed           =-2         # from enum MsoCalloutDropType
	msoCalloutDropTop             =2          # from enum MsoCalloutDropType
	msoCalloutFour                =4          # from enum MsoCalloutType
	msoCalloutMixed               =-2         # from enum MsoCalloutType
	msoCalloutOne                 =1          # from enum MsoCalloutType
	msoCalloutThree               =3          # from enum MsoCalloutType
	msoCalloutTwo                 =2          # from enum MsoCalloutType
	msoCharacterSetArabic         =1          # from enum MsoCharacterSet
	msoCharacterSetCyrillic       =2          # from enum MsoCharacterSet
	msoCharacterSetEnglishWesternEuropeanOtherLatinScript=3          # from enum MsoCharacterSet
	msoCharacterSetGreek          =4          # from enum MsoCharacterSet
	msoCharacterSetHebrew         =5          # from enum MsoCharacterSet
	msoCharacterSetJapanese       =6          # from enum MsoCharacterSet
	msoCharacterSetKorean         =7          # from enum MsoCharacterSet
	msoCharacterSetMultilingualUnicode=8          # from enum MsoCharacterSet
	msoCharacterSetSimplifiedChinese=9          # from enum MsoCharacterSet
	msoCharacterSetThai           =10         # from enum MsoCharacterSet
	msoCharacterSetTraditionalChinese=11         # from enum MsoCharacterSet
	msoCharacterSetVietnamese     =12         # from enum MsoCharacterSet
	msoElementChartFloorNone      =1200       # from enum MsoChartElementType
	msoElementChartFloorShow      =1201       # from enum MsoChartElementType
	msoElementChartTitleAboveChart=2          # from enum MsoChartElementType
	msoElementChartTitleCenteredOverlay=1          # from enum MsoChartElementType
	msoElementChartTitleNone      =0          # from enum MsoChartElementType
	msoElementChartWallNone       =1100       # from enum MsoChartElementType
	msoElementChartWallShow       =1101       # from enum MsoChartElementType
	msoElementDataLabelBestFit    =210        # from enum MsoChartElementType
	msoElementDataLabelBottom     =209        # from enum MsoChartElementType
	msoElementDataLabelCenter     =202        # from enum MsoChartElementType
	msoElementDataLabelInsideBase =204        # from enum MsoChartElementType
	msoElementDataLabelInsideEnd  =203        # from enum MsoChartElementType
	msoElementDataLabelLeft       =206        # from enum MsoChartElementType
	msoElementDataLabelNone       =200        # from enum MsoChartElementType
	msoElementDataLabelOutSideEnd =205        # from enum MsoChartElementType
	msoElementDataLabelRight      =207        # from enum MsoChartElementType
	msoElementDataLabelShow       =201        # from enum MsoChartElementType
	msoElementDataLabelTop        =208        # from enum MsoChartElementType
	msoElementDataTableNone       =500        # from enum MsoChartElementType
	msoElementDataTableShow       =501        # from enum MsoChartElementType
	msoElementDataTableWithLegendKeys=502        # from enum MsoChartElementType
	msoElementErrorBarNone        =700        # from enum MsoChartElementType
	msoElementErrorBarPercentage  =702        # from enum MsoChartElementType
	msoElementErrorBarStandardDeviation=703        # from enum MsoChartElementType
	msoElementErrorBarStandardError=701        # from enum MsoChartElementType
	msoElementLegendBottom        =104        # from enum MsoChartElementType
	msoElementLegendLeft          =103        # from enum MsoChartElementType
	msoElementLegendLeftOverlay   =106        # from enum MsoChartElementType
	msoElementLegendNone          =100        # from enum MsoChartElementType
	msoElementLegendRight         =101        # from enum MsoChartElementType
	msoElementLegendRightOverlay  =105        # from enum MsoChartElementType
	msoElementLegendTop           =102        # from enum MsoChartElementType
	msoElementLineDropHiLoLine    =804        # from enum MsoChartElementType
	msoElementLineDropLine        =801        # from enum MsoChartElementType
	msoElementLineHiLoLine        =802        # from enum MsoChartElementType
	msoElementLineNone            =800        # from enum MsoChartElementType
	msoElementLineSeriesLine      =803        # from enum MsoChartElementType
	msoElementPlotAreaNone        =1000       # from enum MsoChartElementType
	msoElementPlotAreaShow        =1001       # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisBillions=374        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisLogScale=375        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisMillions=373        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisNone=348        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisReverse=351        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisShow=349        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisThousands=372        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisTitleAdjacentToAxis=301        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisTitleBelowAxis=302        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisTitleHorizontal=305        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisTitleNone=300        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisTitleRotated=303        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisTitleVertical=304        # from enum MsoChartElementType
	msoElementPrimaryCategoryAxisWithoutLabels=350        # from enum MsoChartElementType
	msoElementPrimaryCategoryGridLinesMajor=334        # from enum MsoChartElementType
	msoElementPrimaryCategoryGridLinesMinor=333        # from enum MsoChartElementType
	msoElementPrimaryCategoryGridLinesMinorMajor=335        # from enum MsoChartElementType
	msoElementPrimaryCategoryGridLinesNone=332        # from enum MsoChartElementType
	msoElementPrimaryValueAxisBillions=356        # from enum MsoChartElementType
	msoElementPrimaryValueAxisLogScale=357        # from enum MsoChartElementType
	msoElementPrimaryValueAxisMillions=355        # from enum MsoChartElementType
	msoElementPrimaryValueAxisNone=352        # from enum MsoChartElementType
	msoElementPrimaryValueAxisShow=353        # from enum MsoChartElementType
	msoElementPrimaryValueAxisThousands=354        # from enum MsoChartElementType
	msoElementPrimaryValueAxisTitleAdjacentToAxis=306        # from enum MsoChartElementType
	msoElementPrimaryValueAxisTitleBelowAxis=308        # from enum MsoChartElementType
	msoElementPrimaryValueAxisTitleHorizontal=311        # from enum MsoChartElementType
	msoElementPrimaryValueAxisTitleNone=306        # from enum MsoChartElementType
	msoElementPrimaryValueAxisTitleRotated=309        # from enum MsoChartElementType
	msoElementPrimaryValueAxisTitleVertical=310        # from enum MsoChartElementType
	msoElementPrimaryValueGridLinesMajor=330        # from enum MsoChartElementType
	msoElementPrimaryValueGridLinesMinor=329        # from enum MsoChartElementType
	msoElementPrimaryValueGridLinesMinorMajor=331        # from enum MsoChartElementType
	msoElementPrimaryValueGridLinesNone=328        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisBillions=378        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisLogScale=379        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisMillions=377        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisNone=358        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisReverse=361        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisShow=359        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisThousands=376        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisTitleAdjacentToAxis=313        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisTitleBelowAxis=314        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisTitleHorizontal=317        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisTitleNone=312        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisTitleRotated=315        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisTitleVertical=316        # from enum MsoChartElementType
	msoElementSecondaryCategoryAxisWithoutLabels=360        # from enum MsoChartElementType
	msoElementSecondaryCategoryGridLinesMajor=342        # from enum MsoChartElementType
	msoElementSecondaryCategoryGridLinesMinor=341        # from enum MsoChartElementType
	msoElementSecondaryCategoryGridLinesMinorMajor=343        # from enum MsoChartElementType
	msoElementSecondaryCategoryGridLinesNone=340        # from enum MsoChartElementType
	msoElementSecondaryValueAxisBillions=366        # from enum MsoChartElementType
	msoElementSecondaryValueAxisLogScale=367        # from enum MsoChartElementType
	msoElementSecondaryValueAxisMillions=365        # from enum MsoChartElementType
	msoElementSecondaryValueAxisNone=362        # from enum MsoChartElementType
	msoElementSecondaryValueAxisShow=363        # from enum MsoChartElementType
	msoElementSecondaryValueAxisThousands=364        # from enum MsoChartElementType
	msoElementSecondaryValueAxisTitleAdjacentToAxis=319        # from enum MsoChartElementType
	msoElementSecondaryValueAxisTitleBelowAxis=320        # from enum MsoChartElementType
	msoElementSecondaryValueAxisTitleHorizontal=323        # from enum MsoChartElementType
	msoElementSecondaryValueAxisTitleNone=318        # from enum MsoChartElementType
	msoElementSecondaryValueAxisTitleRotated=321        # from enum MsoChartElementType
	msoElementSecondaryValueAxisTitleVertical=322        # from enum MsoChartElementType
	msoElementSecondaryValueGridLinesMajor=338        # from enum MsoChartElementType
	msoElementSecondaryValueGridLinesMinor=337        # from enum MsoChartElementType
	msoElementSecondaryValueGridLinesMinorMajor=339        # from enum MsoChartElementType
	msoElementSecondaryValueGridLinesNone=336        # from enum MsoChartElementType
	msoElementSeriesAxisGridLinesMajor=346        # from enum MsoChartElementType
	msoElementSeriesAxisGridLinesMinor=345        # from enum MsoChartElementType
	msoElementSeriesAxisGridLinesMinorMajor=347        # from enum MsoChartElementType
	msoElementSeriesAxisGridLinesNone=344        # from enum MsoChartElementType
	msoElementSeriesAxisNone      =368        # from enum MsoChartElementType
	msoElementSeriesAxisReverse   =371        # from enum MsoChartElementType
	msoElementSeriesAxisShow      =369        # from enum MsoChartElementType
	msoElementSeriesAxisTitleHorizontal=327        # from enum MsoChartElementType
	msoElementSeriesAxisTitleNone =324        # from enum MsoChartElementType
	msoElementSeriesAxisTitleRotated=325        # from enum MsoChartElementType
	msoElementSeriesAxisTitleVertical=326        # from enum MsoChartElementType
	msoElementSeriesAxisWithoutLabeling=370        # from enum MsoChartElementType
	msoElementTrendlineAddExponential=602        # from enum MsoChartElementType
	msoElementTrendlineAddLinear  =601        # from enum MsoChartElementType
	msoElementTrendlineAddLinearForecast=603        # from enum MsoChartElementType
	msoElementTrendlineAddTwoPeriodMovingAverage=604        # from enum MsoChartElementType
	msoElementTrendlineNone       =600        # from enum MsoChartElementType
	msoElementUpDownBarsNone      =900        # from enum MsoChartElementType
	msoElementUpDownBarsShow      =901        # from enum MsoChartElementType
	msoClipboardFormatHTML        =2          # from enum MsoClipboardFormat
	msoClipboardFormatMixed       =-2         # from enum MsoClipboardFormat
	msoClipboardFormatNative      =1          # from enum MsoClipboardFormat
	msoClipboardFormatPlainText   =4          # from enum MsoClipboardFormat
	msoClipboardFormatRTF         =3          # from enum MsoClipboardFormat
	msoColorTypeCMS               =4          # from enum MsoColorType
	msoColorTypeCMYK              =3          # from enum MsoColorType
	msoColorTypeInk               =5          # from enum MsoColorType
	msoColorTypeMixed             =-2         # from enum MsoColorType
	msoColorTypeRGB               =1          # from enum MsoColorType
	msoColorTypeScheme            =2          # from enum MsoColorType
	msoComboLabel                 =1          # from enum MsoComboStyle
	msoComboNormal                =0          # from enum MsoComboStyle
	msoCommandBarButtonHyperlinkInsertPicture=2          # from enum MsoCommandBarButtonHyperlinkType
	msoCommandBarButtonHyperlinkNone=0          # from enum MsoCommandBarButtonHyperlinkType
	msoCommandBarButtonHyperlinkOpen=1          # from enum MsoCommandBarButtonHyperlinkType
	msoConditionAnyNumberBetween  =34         # from enum MsoCondition
	msoConditionAnytime           =25         # from enum MsoCondition
	msoConditionAnytimeBetween    =26         # from enum MsoCondition
	msoConditionAtLeast           =36         # from enum MsoCondition
	msoConditionAtMost            =35         # from enum MsoCondition
	msoConditionBeginsWith        =11         # from enum MsoCondition
	msoConditionDoesNotEqual      =33         # from enum MsoCondition
	msoConditionEndsWith          =12         # from enum MsoCondition
	msoConditionEquals            =32         # from enum MsoCondition
	msoConditionEqualsCompleted   =66         # from enum MsoCondition
	msoConditionEqualsDeferred    =68         # from enum MsoCondition
	msoConditionEqualsHigh        =60         # from enum MsoCondition
	msoConditionEqualsInProgress  =65         # from enum MsoCondition
	msoConditionEqualsLow         =58         # from enum MsoCondition
	msoConditionEqualsNormal      =59         # from enum MsoCondition
	msoConditionEqualsNotStarted  =64         # from enum MsoCondition
	msoConditionEqualsWaitingForSomeoneElse=67         # from enum MsoCondition
	msoConditionFileTypeAllFiles  =1          # from enum MsoCondition
	msoConditionFileTypeBinders   =6          # from enum MsoCondition
	msoConditionFileTypeCalendarItem=45         # from enum MsoCondition
	msoConditionFileTypeContactItem=46         # from enum MsoCondition
	msoConditionFileTypeDataConnectionFiles=51         # from enum MsoCondition
	msoConditionFileTypeDatabases =7          # from enum MsoCondition
	msoConditionFileTypeDesignerFiles=56         # from enum MsoCondition
	msoConditionFileTypeDocumentImagingFiles=54         # from enum MsoCondition
	msoConditionFileTypeExcelWorkbooks=4          # from enum MsoCondition
	msoConditionFileTypeJournalItem=48         # from enum MsoCondition
	msoConditionFileTypeMailItem  =44         # from enum MsoCondition
	msoConditionFileTypeNoteItem  =47         # from enum MsoCondition
	msoConditionFileTypeOfficeFiles=2          # from enum MsoCondition
	msoConditionFileTypeOutlookItems=43         # from enum MsoCondition
	msoConditionFileTypePhotoDrawFiles=50         # from enum MsoCondition
	msoConditionFileTypePowerPointPresentations=5          # from enum MsoCondition
	msoConditionFileTypeProjectFiles=53         # from enum MsoCondition
	msoConditionFileTypePublisherFiles=52         # from enum MsoCondition
	msoConditionFileTypeTaskItem  =49         # from enum MsoCondition
	msoConditionFileTypeTemplates =8          # from enum MsoCondition
	msoConditionFileTypeVisioFiles=55         # from enum MsoCondition
	msoConditionFileTypeWebPages  =57         # from enum MsoCondition
	msoConditionFileTypeWordDocuments=3          # from enum MsoCondition
	msoConditionFreeText          =42         # from enum MsoCondition
	msoConditionInTheLast         =31         # from enum MsoCondition
	msoConditionInTheNext         =30         # from enum MsoCondition
	msoConditionIncludes          =9          # from enum MsoCondition
	msoConditionIncludesFormsOf   =41         # from enum MsoCondition
	msoConditionIncludesNearEachOther=13         # from enum MsoCondition
	msoConditionIncludesPhrase    =10         # from enum MsoCondition
	msoConditionIsExactly         =14         # from enum MsoCondition
	msoConditionIsNo              =40         # from enum MsoCondition
	msoConditionIsNot             =15         # from enum MsoCondition
	msoConditionIsYes             =39         # from enum MsoCondition
	msoConditionLastMonth         =22         # from enum MsoCondition
	msoConditionLastWeek          =19         # from enum MsoCondition
	msoConditionLessThan          =38         # from enum MsoCondition
	msoConditionMoreThan          =37         # from enum MsoCondition
	msoConditionNextMonth         =24         # from enum MsoCondition
	msoConditionNextWeek          =21         # from enum MsoCondition
	msoConditionNotEqualToCompleted=71         # from enum MsoCondition
	msoConditionNotEqualToDeferred=73         # from enum MsoCondition
	msoConditionNotEqualToHigh    =63         # from enum MsoCondition
	msoConditionNotEqualToInProgress=70         # from enum MsoCondition
	msoConditionNotEqualToLow     =61         # from enum MsoCondition
	msoConditionNotEqualToNormal  =62         # from enum MsoCondition
	msoConditionNotEqualToNotStarted=69         # from enum MsoCondition
	msoConditionNotEqualToWaitingForSomeoneElse=72         # from enum MsoCondition
	msoConditionOn                =27         # from enum MsoCondition
	msoConditionOnOrAfter         =28         # from enum MsoCondition
	msoConditionOnOrBefore        =29         # from enum MsoCondition
	msoConditionThisMonth         =23         # from enum MsoCondition
	msoConditionThisWeek          =20         # from enum MsoCondition
	msoConditionToday             =17         # from enum MsoCondition
	msoConditionTomorrow          =18         # from enum MsoCondition
	msoConditionYesterday         =16         # from enum MsoCondition
	msoConnectorAnd               =1          # from enum MsoConnector
	msoConnectorOr                =2          # from enum MsoConnector
	msoConnectorCurve             =3          # from enum MsoConnectorType
	msoConnectorElbow             =2          # from enum MsoConnectorType
	msoConnectorStraight          =1          # from enum MsoConnectorType
	msoConnectorTypeMixed         =-2         # from enum MsoConnectorType
	msoControlOLEUsageBoth        =3          # from enum MsoControlOLEUsage
	msoControlOLEUsageClient      =2          # from enum MsoControlOLEUsage
	msoControlOLEUsageNeither     =0          # from enum MsoControlOLEUsage
	msoControlOLEUsageServer      =1          # from enum MsoControlOLEUsage
	msoControlActiveX             =22         # from enum MsoControlType
	msoControlAutoCompleteCombo   =26         # from enum MsoControlType
	msoControlButton              =1          # from enum MsoControlType
	msoControlButtonDropdown      =5          # from enum MsoControlType
	msoControlButtonPopup         =12         # from enum MsoControlType
	msoControlComboBox            =4          # from enum MsoControlType
	msoControlCustom              =0          # from enum MsoControlType
	msoControlDropdown            =3          # from enum MsoControlType
	msoControlEdit                =2          # from enum MsoControlType
	msoControlExpandingGrid       =16         # from enum MsoControlType
	msoControlGauge               =19         # from enum MsoControlType
	msoControlGenericDropdown     =8          # from enum MsoControlType
	msoControlGraphicCombo        =20         # from enum MsoControlType
	msoControlGraphicDropdown     =9          # from enum MsoControlType
	msoControlGraphicPopup        =11         # from enum MsoControlType
	msoControlGrid                =18         # from enum MsoControlType
	msoControlLabel               =15         # from enum MsoControlType
	msoControlLabelEx             =24         # from enum MsoControlType
	msoControlOCXDropdown         =7          # from enum MsoControlType
	msoControlPane                =21         # from enum MsoControlType
	msoControlPopup               =10         # from enum MsoControlType
	msoControlSpinner             =23         # from enum MsoControlType
	msoControlSplitButtonMRUPopup =14         # from enum MsoControlType
	msoControlSplitButtonPopup    =13         # from enum MsoControlType
	msoControlSplitDropdown       =6          # from enum MsoControlType
	msoControlSplitExpandingGrid  =17         # from enum MsoControlType
	msoControlWorkPane            =25         # from enum MsoControlType
	msoCustomXMLNodeAttribute     =2          # from enum MsoCustomXMLNodeType
	msoCustomXMLNodeCData         =4          # from enum MsoCustomXMLNodeType
	msoCustomXMLNodeComment       =8          # from enum MsoCustomXMLNodeType
	msoCustomXMLNodeDocument      =9          # from enum MsoCustomXMLNodeType
	msoCustomXMLNodeElement       =1          # from enum MsoCustomXMLNodeType
	msoCustomXMLNodeProcessingInstruction=7          # from enum MsoCustomXMLNodeType
	msoCustomXMLNodeText          =3          # from enum MsoCustomXMLNodeType
	msoCustomXMLValidationErrorAutomaticallyCleared=1          # from enum MsoCustomXMLValidationErrorType
	msoCustomXMLValidationErrorManual=2          # from enum MsoCustomXMLValidationErrorType
	msoCustomXMLValidationErrorSchemaGenerated=0          # from enum MsoCustomXMLValidationErrorType
	msoDateTimeFigureOut          =14         # from enum MsoDateTimeFormat
	msoDateTimeFormatMixed        =-2         # from enum MsoDateTimeFormat
	msoDateTimeHmm                =10         # from enum MsoDateTimeFormat
	msoDateTimeHmmss              =11         # from enum MsoDateTimeFormat
	msoDateTimeMMMMdyyyy          =4          # from enum MsoDateTimeFormat
	msoDateTimeMMMMyy             =6          # from enum MsoDateTimeFormat
	msoDateTimeMMddyyHmm          =8          # from enum MsoDateTimeFormat
	msoDateTimeMMddyyhmmAMPM      =9          # from enum MsoDateTimeFormat
	msoDateTimeMMyy               =7          # from enum MsoDateTimeFormat
	msoDateTimeMdyy               =1          # from enum MsoDateTimeFormat
	msoDateTimedMMMMyyyy          =3          # from enum MsoDateTimeFormat
	msoDateTimedMMMyy             =5          # from enum MsoDateTimeFormat
	msoDateTimeddddMMMMddyyyy     =2          # from enum MsoDateTimeFormat
	msoDateTimehmmAMPM            =12         # from enum MsoDateTimeFormat
	msoDateTimehmmssAMPM          =13         # from enum MsoDateTimeFormat
	msoDiagramAssistant           =2          # from enum MsoDiagramNodeType
	msoDiagramNode                =1          # from enum MsoDiagramNodeType
	msoDiagramCycle               =2          # from enum MsoDiagramType
	msoDiagramMixed               =-2         # from enum MsoDiagramType
	msoDiagramOrgChart            =1          # from enum MsoDiagramType
	msoDiagramPyramid             =4          # from enum MsoDiagramType
	msoDiagramRadial              =3          # from enum MsoDiagramType
	msoDiagramTarget              =6          # from enum MsoDiagramType
	msoDiagramVenn                =5          # from enum MsoDiagramType
	msoDistributeHorizontally     =0          # from enum MsoDistributeCmd
	msoDistributeVertically       =1          # from enum MsoDistributeCmd
	msoDocInspectorStatusDocOk    =0          # from enum MsoDocInspectorStatus
	msoDocInspectorStatusError    =2          # from enum MsoDocInspectorStatus
	msoDocInspectorStatusIssueFound=1          # from enum MsoDocInspectorStatus
	msoPropertyTypeBoolean        =2          # from enum MsoDocProperties
	msoPropertyTypeDate           =3          # from enum MsoDocProperties
	msoPropertyTypeFloat          =5          # from enum MsoDocProperties
	msoPropertyTypeNumber         =1          # from enum MsoDocProperties
	msoPropertyTypeString         =4          # from enum MsoDocProperties
	msoEditingAuto                =0          # from enum MsoEditingType
	msoEditingCorner              =1          # from enum MsoEditingType
	msoEditingSmooth              =2          # from enum MsoEditingType
	msoEditingSymmetric           =3          # from enum MsoEditingType
	msoEncodingArabic             =1256       # from enum MsoEncoding
	msoEncodingArabicASMO         =708        # from enum MsoEncoding
	msoEncodingArabicAutoDetect   =51256      # from enum MsoEncoding
	msoEncodingArabicTransparentASMO=720        # from enum MsoEncoding
	msoEncodingAutoDetect         =50001      # from enum MsoEncoding
	msoEncodingBaltic             =1257       # from enum MsoEncoding
	msoEncodingCentralEuropean    =1250       # from enum MsoEncoding
	msoEncodingCyrillic           =1251       # from enum MsoEncoding
	msoEncodingCyrillicAutoDetect =51251      # from enum MsoEncoding
	msoEncodingEBCDICArabic       =20420      # from enum MsoEncoding
	msoEncodingEBCDICDenmarkNorway=20277      # from enum MsoEncoding
	msoEncodingEBCDICFinlandSweden=20278      # from enum MsoEncoding
	msoEncodingEBCDICFrance       =20297      # from enum MsoEncoding
	msoEncodingEBCDICGermany      =20273      # from enum MsoEncoding
	msoEncodingEBCDICGreek        =20423      # from enum MsoEncoding
	msoEncodingEBCDICGreekModern  =875        # from enum MsoEncoding
	msoEncodingEBCDICHebrew       =20424      # from enum MsoEncoding
	msoEncodingEBCDICIcelandic    =20871      # from enum MsoEncoding
	msoEncodingEBCDICInternational=500        # from enum MsoEncoding
	msoEncodingEBCDICItaly        =20280      # from enum MsoEncoding
	msoEncodingEBCDICJapaneseKatakanaExtended=20290      # from enum MsoEncoding
	msoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese=50930      # from enum MsoEncoding
	msoEncodingEBCDICJapaneseLatinExtendedAndJapanese=50939      # from enum MsoEncoding
	msoEncodingEBCDICKoreanExtended=20833      # from enum MsoEncoding
	msoEncodingEBCDICKoreanExtendedAndKorean=50933      # from enum MsoEncoding
	msoEncodingEBCDICLatinAmericaSpain=20284      # from enum MsoEncoding
	msoEncodingEBCDICMultilingualROECELatin2=870        # from enum MsoEncoding
	msoEncodingEBCDICRussian      =20880      # from enum MsoEncoding
	msoEncodingEBCDICSerbianBulgarian=21025      # from enum MsoEncoding
	msoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese=50935      # from enum MsoEncoding
	msoEncodingEBCDICThai         =20838      # from enum MsoEncoding
	msoEncodingEBCDICTurkish      =20905      # from enum MsoEncoding
	msoEncodingEBCDICTurkishLatin5=1026       # from enum MsoEncoding
	msoEncodingEBCDICUSCanada     =37         # from enum MsoEncoding
	msoEncodingEBCDICUSCanadaAndJapanese=50931      # from enum MsoEncoding
	msoEncodingEBCDICUSCanadaAndTraditionalChinese=50937      # from enum MsoEncoding
	msoEncodingEBCDICUnitedKingdom=20285      # from enum MsoEncoding
	msoEncodingEUCChineseSimplifiedChinese=51936      # from enum MsoEncoding
	msoEncodingEUCJapanese        =51932      # from enum MsoEncoding
	msoEncodingEUCKorean          =51949      # from enum MsoEncoding
	msoEncodingEUCTaiwaneseTraditionalChinese=51950      # from enum MsoEncoding
	msoEncodingEuropa3            =29001      # from enum MsoEncoding
	msoEncodingExtAlphaLowercase  =21027      # from enum MsoEncoding
	msoEncodingGreek              =1253       # from enum MsoEncoding
	msoEncodingGreekAutoDetect    =51253      # from enum MsoEncoding
	msoEncodingHZGBSimplifiedChinese=52936      # from enum MsoEncoding
	msoEncodingHebrew             =1255       # from enum MsoEncoding
	msoEncodingIA5German          =20106      # from enum MsoEncoding
	msoEncodingIA5IRV             =20105      # from enum MsoEncoding
	msoEncodingIA5Norwegian       =20108      # from enum MsoEncoding
	msoEncodingIA5Swedish         =20107      # from enum MsoEncoding
	msoEncodingISCIIAssamese      =57006      # from enum MsoEncoding
	msoEncodingISCIIBengali       =57003      # from enum MsoEncoding
	msoEncodingISCIIDevanagari    =57002      # from enum MsoEncoding
	msoEncodingISCIIGujarati      =57010      # from enum MsoEncoding
	msoEncodingISCIIKannada       =57008      # from enum MsoEncoding
	msoEncodingISCIIMalayalam     =57009      # from enum MsoEncoding
	msoEncodingISCIIOriya         =57007      # from enum MsoEncoding
	msoEncodingISCIIPunjabi       =57011      # from enum MsoEncoding
	msoEncodingISCIITamil         =57004      # from enum MsoEncoding
	msoEncodingISCIITelugu        =57005      # from enum MsoEncoding
	msoEncodingISO2022CNSimplifiedChinese=50229      # from enum MsoEncoding
	msoEncodingISO2022CNTraditionalChinese=50227      # from enum MsoEncoding
	msoEncodingISO2022JPJISX02011989=50222      # from enum MsoEncoding
	msoEncodingISO2022JPJISX02021984=50221      # from enum MsoEncoding
	msoEncodingISO2022JPNoHalfwidthKatakana=50220      # from enum MsoEncoding
	msoEncodingISO2022KR          =50225      # from enum MsoEncoding
	msoEncodingISO6937NonSpacingAccent=20269      # from enum MsoEncoding
	msoEncodingISO885915Latin9    =28605      # from enum MsoEncoding
	msoEncodingISO88591Latin1     =28591      # from enum MsoEncoding
	msoEncodingISO88592CentralEurope=28592      # from enum MsoEncoding
	msoEncodingISO88593Latin3     =28593      # from enum MsoEncoding
	msoEncodingISO88594Baltic     =28594      # from enum MsoEncoding
	msoEncodingISO88595Cyrillic   =28595      # from enum MsoEncoding
	msoEncodingISO88596Arabic     =28596      # from enum MsoEncoding
	msoEncodingISO88597Greek      =28597      # from enum MsoEncoding
	msoEncodingISO88598Hebrew     =28598      # from enum MsoEncoding
	msoEncodingISO88598HebrewLogical=38598      # from enum MsoEncoding
	msoEncodingISO88599Turkish    =28599      # from enum MsoEncoding
	msoEncodingJapaneseAutoDetect =50932      # from enum MsoEncoding
	msoEncodingJapaneseShiftJIS   =932        # from enum MsoEncoding
	msoEncodingKOI8R              =20866      # from enum MsoEncoding
	msoEncodingKOI8U              =21866      # from enum MsoEncoding
	msoEncodingKorean             =949        # from enum MsoEncoding
	msoEncodingKoreanAutoDetect   =50949      # from enum MsoEncoding
	msoEncodingKoreanJohab        =1361       # from enum MsoEncoding
	msoEncodingMacArabic          =10004      # from enum MsoEncoding
	msoEncodingMacCroatia         =10082      # from enum MsoEncoding
	msoEncodingMacCyrillic        =10007      # from enum MsoEncoding
	msoEncodingMacGreek1          =10006      # from enum MsoEncoding
	msoEncodingMacHebrew          =10005      # from enum MsoEncoding
	msoEncodingMacIcelandic       =10079      # from enum MsoEncoding
	msoEncodingMacJapanese        =10001      # from enum MsoEncoding
	msoEncodingMacKorean          =10003      # from enum MsoEncoding
	msoEncodingMacLatin2          =10029      # from enum MsoEncoding
	msoEncodingMacRoman           =10000      # from enum MsoEncoding
	msoEncodingMacRomania         =10010      # from enum MsoEncoding
	msoEncodingMacSimplifiedChineseGB2312=10008      # from enum MsoEncoding
	msoEncodingMacTraditionalChineseBig5=10002      # from enum MsoEncoding
	msoEncodingMacTurkish         =10081      # from enum MsoEncoding
	msoEncodingMacUkraine         =10017      # from enum MsoEncoding
	msoEncodingOEMArabic          =864        # from enum MsoEncoding
	msoEncodingOEMBaltic          =775        # from enum MsoEncoding
	msoEncodingOEMCanadianFrench  =863        # from enum MsoEncoding
	msoEncodingOEMCyrillic        =855        # from enum MsoEncoding
	msoEncodingOEMCyrillicII      =866        # from enum MsoEncoding
	msoEncodingOEMGreek437G       =737        # from enum MsoEncoding
	msoEncodingOEMHebrew          =862        # from enum MsoEncoding
	msoEncodingOEMIcelandic       =861        # from enum MsoEncoding
	msoEncodingOEMModernGreek     =869        # from enum MsoEncoding
	msoEncodingOEMMultilingualLatinI=850        # from enum MsoEncoding
	msoEncodingOEMMultilingualLatinII=852        # from enum MsoEncoding
	msoEncodingOEMNordic          =865        # from enum MsoEncoding
	msoEncodingOEMPortuguese      =860        # from enum MsoEncoding
	msoEncodingOEMTurkish         =857        # from enum MsoEncoding
	msoEncodingOEMUnitedStates    =437        # from enum MsoEncoding
	msoEncodingSimplifiedChineseAutoDetect=50936      # from enum MsoEncoding
	msoEncodingSimplifiedChineseGB18030=54936      # from enum MsoEncoding
	msoEncodingSimplifiedChineseGBK=936        # from enum MsoEncoding
	msoEncodingT61                =20261      # from enum MsoEncoding
	msoEncodingTaiwanCNS          =20000      # from enum MsoEncoding
	msoEncodingTaiwanEten         =20002      # from enum MsoEncoding
	msoEncodingTaiwanIBM5550      =20003      # from enum MsoEncoding
	msoEncodingTaiwanTCA          =20001      # from enum MsoEncoding
	msoEncodingTaiwanTeleText     =20004      # from enum MsoEncoding
	msoEncodingTaiwanWang         =20005      # from enum MsoEncoding
	msoEncodingThai               =874        # from enum MsoEncoding
	msoEncodingTraditionalChineseAutoDetect=50950      # from enum MsoEncoding
	msoEncodingTraditionalChineseBig5=950        # from enum MsoEncoding
	msoEncodingTurkish            =1254       # from enum MsoEncoding
	msoEncodingUSASCII            =20127      # from enum MsoEncoding
	msoEncodingUTF7               =65000      # from enum MsoEncoding
	msoEncodingUTF8               =65001      # from enum MsoEncoding
	msoEncodingUnicodeBigEndian   =1201       # from enum MsoEncoding
	msoEncodingUnicodeLittleEndian=1200       # from enum MsoEncoding
	msoEncodingVietnamese         =1258       # from enum MsoEncoding
	msoEncodingWestern            =1252       # from enum MsoEncoding
	msoMethodGet                  =0          # from enum MsoExtraInfoMethod
	msoMethodPost                 =1          # from enum MsoExtraInfoMethod
	msoExtrusionColorAutomatic    =1          # from enum MsoExtrusionColorType
	msoExtrusionColorCustom       =2          # from enum MsoExtrusionColorType
	msoExtrusionColorTypeMixed    =-2         # from enum MsoExtrusionColorType
	MsoFarEastLineBreakLanguageJapanese=1041       # from enum MsoFarEastLineBreakLanguageID
	MsoFarEastLineBreakLanguageKorean=1042       # from enum MsoFarEastLineBreakLanguageID
	MsoFarEastLineBreakLanguageSimplifiedChinese=2052       # from enum MsoFarEastLineBreakLanguageID
	MsoFarEastLineBreakLanguageTraditionalChinese=1028       # from enum MsoFarEastLineBreakLanguageID
	msoFeatureInstallNone         =0          # from enum MsoFeatureInstall
	msoFeatureInstallOnDemand     =1          # from enum MsoFeatureInstall
	msoFeatureInstallOnDemandWithUI=2          # from enum MsoFeatureInstall
	msoFileDialogFilePicker       =3          # from enum MsoFileDialogType
	msoFileDialogFolderPicker     =4          # from enum MsoFileDialogType
	msoFileDialogOpen             =1          # from enum MsoFileDialogType
	msoFileDialogSaveAs           =2          # from enum MsoFileDialogType
	msoFileDialogViewDetails      =2          # from enum MsoFileDialogView
	msoFileDialogViewLargeIcons   =6          # from enum MsoFileDialogView
	msoFileDialogViewList         =1          # from enum MsoFileDialogView
	msoFileDialogViewPreview      =4          # from enum MsoFileDialogView
	msoFileDialogViewProperties   =3          # from enum MsoFileDialogView
	msoFileDialogViewSmallIcons   =7          # from enum MsoFileDialogView
	msoFileDialogViewThumbnail    =5          # from enum MsoFileDialogView
	msoFileDialogViewTiles        =9          # from enum MsoFileDialogView
	msoFileDialogViewWebView      =8          # from enum MsoFileDialogView
	msoListbyName                 =1          # from enum MsoFileFindListBy
	msoListbyTitle                =2          # from enum MsoFileFindListBy
	msoOptionsAdd                 =2          # from enum MsoFileFindOptions
	msoOptionsNew                 =1          # from enum MsoFileFindOptions
	msoOptionsWithin              =3          # from enum MsoFileFindOptions
	msoFileFindSortbyAuthor       =1          # from enum MsoFileFindSortBy
	msoFileFindSortbyDateCreated  =2          # from enum MsoFileFindSortBy
	msoFileFindSortbyDateSaved    =4          # from enum MsoFileFindSortBy
	msoFileFindSortbyFileName     =5          # from enum MsoFileFindSortBy
	msoFileFindSortbyLastSavedBy  =3          # from enum MsoFileFindSortBy
	msoFileFindSortbySize         =6          # from enum MsoFileFindSortBy
	msoFileFindSortbyTitle        =7          # from enum MsoFileFindSortBy
	msoViewFileInfo               =1          # from enum MsoFileFindView
	msoViewPreview                =2          # from enum MsoFileFindView
	msoViewSummaryInfo            =3          # from enum MsoFileFindView
	msoCreateNewFile              =1          # from enum MsoFileNewAction
	msoEditFile                   =0          # from enum MsoFileNewAction
	msoOpenFile                   =2          # from enum MsoFileNewAction
	msoBottomSection              =4          # from enum MsoFileNewSection
	msoNew                        =1          # from enum MsoFileNewSection
	msoNewfromExistingFile        =2          # from enum MsoFileNewSection
	msoNewfromTemplate            =3          # from enum MsoFileNewSection
	msoOpenDocument               =0          # from enum MsoFileNewSection
	msoFileTypeAllFiles           =1          # from enum MsoFileType
	msoFileTypeBinders            =6          # from enum MsoFileType
	msoFileTypeCalendarItem       =11         # from enum MsoFileType
	msoFileTypeContactItem        =12         # from enum MsoFileType
	msoFileTypeDataConnectionFiles=17         # from enum MsoFileType
	msoFileTypeDatabases          =7          # from enum MsoFileType
	msoFileTypeDesignerFiles      =22         # from enum MsoFileType
	msoFileTypeDocumentImagingFiles=20         # from enum MsoFileType
	msoFileTypeExcelWorkbooks     =4          # from enum MsoFileType
	msoFileTypeJournalItem        =14         # from enum MsoFileType
	msoFileTypeMailItem           =10         # from enum MsoFileType
	msoFileTypeNoteItem           =13         # from enum MsoFileType
	msoFileTypeOfficeFiles        =2          # from enum MsoFileType
	msoFileTypeOutlookItems       =9          # from enum MsoFileType
	msoFileTypePhotoDrawFiles     =16         # from enum MsoFileType
	msoFileTypePowerPointPresentations=5          # from enum MsoFileType
	msoFileTypeProjectFiles       =19         # from enum MsoFileType
	msoFileTypePublisherFiles     =18         # from enum MsoFileType
	msoFileTypeTaskItem           =15         # from enum MsoFileType
	msoFileTypeTemplates          =8          # from enum MsoFileType
	msoFileTypeVisioFiles         =21         # from enum MsoFileType
	msoFileTypeWebPages           =23         # from enum MsoFileType
	msoFileTypeWordDocuments      =3          # from enum MsoFileType
	msoFillBackground             =5          # from enum MsoFillType
	msoFillGradient               =3          # from enum MsoFillType
	msoFillMixed                  =-2         # from enum MsoFillType
	msoFillPatterned              =2          # from enum MsoFillType
	msoFillPicture                =6          # from enum MsoFillType
	msoFillSolid                  =1          # from enum MsoFillType
	msoFillTextured               =4          # from enum MsoFillType
	msoFilterComparisonContains   =8          # from enum MsoFilterComparison
	msoFilterComparisonEqual      =0          # from enum MsoFilterComparison
	msoFilterComparisonGreaterThan=3          # from enum MsoFilterComparison
	msoFilterComparisonGreaterThanEqual=5          # from enum MsoFilterComparison
	msoFilterComparisonIsBlank    =6          # from enum MsoFilterComparison
	msoFilterComparisonIsNotBlank =7          # from enum MsoFilterComparison
	msoFilterComparisonLessThan   =2          # from enum MsoFilterComparison
	msoFilterComparisonLessThanEqual=4          # from enum MsoFilterComparison
	msoFilterComparisonNotContains=9          # from enum MsoFilterComparison
	msoFilterComparisonNotEqual   =1          # from enum MsoFilterComparison
	msoFilterConjunctionAnd       =0          # from enum MsoFilterConjunction
	msoFilterConjunctionOr        =1          # from enum MsoFilterConjunction
	msoFlipHorizontal             =0          # from enum MsoFlipCmd
	msoFlipVertical               =1          # from enum MsoFlipCmd
	msoThemeComplexScript         =2          # from enum MsoFontLanguageIndex
	msoThemeEastAsian             =3          # from enum MsoFontLanguageIndex
	msoThemeLatin                 =1          # from enum MsoFontLanguageIndex
	msoGradientColorMixed         =-2         # from enum MsoGradientColorType
	msoGradientMultiColor         =4          # from enum MsoGradientColorType
	msoGradientOneColor           =1          # from enum MsoGradientColorType
	msoGradientPresetColors       =3          # from enum MsoGradientColorType
	msoGradientTwoColors          =2          # from enum MsoGradientColorType
	msoGradientDiagonalDown       =4          # from enum MsoGradientStyle
	msoGradientDiagonalUp         =3          # from enum MsoGradientStyle
	msoGradientFromCenter         =7          # from enum MsoGradientStyle
	msoGradientFromCorner         =5          # from enum MsoGradientStyle
	msoGradientFromTitle          =6          # from enum MsoGradientStyle
	msoGradientHorizontal         =1          # from enum MsoGradientStyle
	msoGradientMixed              =-2         # from enum MsoGradientStyle
	msoGradientVertical           =2          # from enum MsoGradientStyle
	msoHTMLProjectOpenSourceView  =1          # from enum MsoHTMLProjectOpen
	msoHTMLProjectOpenTextView    =2          # from enum MsoHTMLProjectOpen
	msoHTMLProjectStateDocumentLocked=1          # from enum MsoHTMLProjectState
	msoHTMLProjectStateDocumentProjectUnlocked=3          # from enum MsoHTMLProjectState
	msoHTMLProjectStateProjectLocked=2          # from enum MsoHTMLProjectState
	msoAnchorCenter               =2          # from enum MsoHorizontalAnchor
	msoAnchorNone                 =1          # from enum MsoHorizontalAnchor
	msoHorizontalAnchorMixed      =-2         # from enum MsoHorizontalAnchor
	msoHyperlinkInlineShape       =2          # from enum MsoHyperlinkType
	msoHyperlinkRange             =0          # from enum MsoHyperlinkType
	msoHyperlinkShape             =1          # from enum MsoHyperlinkType
	msoIconAlert                  =2          # from enum MsoIconType
	msoIconAlertCritical          =7          # from enum MsoIconType
	msoIconAlertInfo              =4          # from enum MsoIconType
	msoIconAlertQuery             =6          # from enum MsoIconType
	msoIconAlertWarning           =5          # from enum MsoIconType
	msoIconNone                   =0          # from enum MsoIconType
	msoIconTip                    =3          # from enum MsoIconType
	msoLanguageIDAfrikaans        =1078       # from enum MsoLanguageID
	msoLanguageIDAlbanian         =1052       # from enum MsoLanguageID
	msoLanguageIDAmharic          =1118       # from enum MsoLanguageID
	msoLanguageIDArabic           =1025       # from enum MsoLanguageID
	msoLanguageIDArabicAlgeria    =5121       # from enum MsoLanguageID
	msoLanguageIDArabicBahrain    =15361      # from enum MsoLanguageID
	msoLanguageIDArabicEgypt      =3073       # from enum MsoLanguageID
	msoLanguageIDArabicIraq       =2049       # from enum MsoLanguageID
	msoLanguageIDArabicJordan     =11265      # from enum MsoLanguageID
	msoLanguageIDArabicKuwait     =13313      # from enum MsoLanguageID
	msoLanguageIDArabicLebanon    =12289      # from enum MsoLanguageID
	msoLanguageIDArabicLibya      =4097       # from enum MsoLanguageID
	msoLanguageIDArabicMorocco    =6145       # from enum MsoLanguageID
	msoLanguageIDArabicOman       =8193       # from enum MsoLanguageID
	msoLanguageIDArabicQatar      =16385      # from enum MsoLanguageID
	msoLanguageIDArabicSyria      =10241      # from enum MsoLanguageID
	msoLanguageIDArabicTunisia    =7169       # from enum MsoLanguageID
	msoLanguageIDArabicUAE        =14337      # from enum MsoLanguageID
	msoLanguageIDArabicYemen      =9217       # from enum MsoLanguageID
	msoLanguageIDArmenian         =1067       # from enum MsoLanguageID
	msoLanguageIDAssamese         =1101       # from enum MsoLanguageID
	msoLanguageIDAzeriCyrillic    =2092       # from enum MsoLanguageID
	msoLanguageIDAzeriLatin       =1068       # from enum MsoLanguageID
	msoLanguageIDBasque           =1069       # from enum MsoLanguageID
	msoLanguageIDBelgianDutch     =2067       # from enum MsoLanguageID
	msoLanguageIDBelgianFrench    =2060       # from enum MsoLanguageID
	msoLanguageIDBengali          =1093       # from enum MsoLanguageID
	msoLanguageIDBosnian          =4122       # from enum MsoLanguageID
	msoLanguageIDBosnianBosniaHerzegovinaCyrillic=8218       # from enum MsoLanguageID
	msoLanguageIDBosnianBosniaHerzegovinaLatin=5146       # from enum MsoLanguageID
	msoLanguageIDBrazilianPortuguese=1046       # from enum MsoLanguageID
	msoLanguageIDBulgarian        =1026       # from enum MsoLanguageID
	msoLanguageIDBurmese          =1109       # from enum MsoLanguageID
	msoLanguageIDByelorussian     =1059       # from enum MsoLanguageID
	msoLanguageIDCatalan          =1027       # from enum MsoLanguageID
	msoLanguageIDCherokee         =1116       # from enum MsoLanguageID
	msoLanguageIDChineseHongKongSAR=3076       # from enum MsoLanguageID
	msoLanguageIDChineseMacaoSAR  =5124       # from enum MsoLanguageID
	msoLanguageIDChineseSingapore =4100       # from enum MsoLanguageID
	msoLanguageIDCroatian         =1050       # from enum MsoLanguageID
	msoLanguageIDCzech            =1029       # from enum MsoLanguageID
	msoLanguageIDDanish           =1030       # from enum MsoLanguageID
	msoLanguageIDDivehi           =1125       # from enum MsoLanguageID
	msoLanguageIDDutch            =1043       # from enum MsoLanguageID
	msoLanguageIDDzongkhaBhutan   =2129       # from enum MsoLanguageID
	msoLanguageIDEdo              =1126       # from enum MsoLanguageID
	msoLanguageIDEnglishAUS       =3081       # from enum MsoLanguageID
	msoLanguageIDEnglishBelize    =10249      # from enum MsoLanguageID
	msoLanguageIDEnglishCanadian  =4105       # from enum MsoLanguageID
	msoLanguageIDEnglishCaribbean =9225       # from enum MsoLanguageID
	msoLanguageIDEnglishIndonesia =14345      # from enum MsoLanguageID
	msoLanguageIDEnglishIreland   =6153       # from enum MsoLanguageID
	msoLanguageIDEnglishJamaica   =8201       # from enum MsoLanguageID
	msoLanguageIDEnglishNewZealand=5129       # from enum MsoLanguageID
	msoLanguageIDEnglishPhilippines=13321      # from enum MsoLanguageID
	msoLanguageIDEnglishSouthAfrica=7177       # from enum MsoLanguageID
	msoLanguageIDEnglishTrinidadTobago=11273      # from enum MsoLanguageID
	msoLanguageIDEnglishUK        =2057       # from enum MsoLanguageID
	msoLanguageIDEnglishUS        =1033       # from enum MsoLanguageID
	msoLanguageIDEnglishZimbabwe  =12297      # from enum MsoLanguageID
	msoLanguageIDEstonian         =1061       # from enum MsoLanguageID
	msoLanguageIDFaeroese         =1080       # from enum MsoLanguageID
	msoLanguageIDFarsi            =1065       # from enum MsoLanguageID
	msoLanguageIDFilipino         =1124       # from enum MsoLanguageID
	msoLanguageIDFinnish          =1035       # from enum MsoLanguageID
	msoLanguageIDFrench           =1036       # from enum MsoLanguageID
	msoLanguageIDFrenchCameroon   =11276      # from enum MsoLanguageID
	msoLanguageIDFrenchCanadian   =3084       # from enum MsoLanguageID
	msoLanguageIDFrenchCongoDRC   =9228       # from enum MsoLanguageID
	msoLanguageIDFrenchCotedIvoire=12300      # from enum MsoLanguageID
	msoLanguageIDFrenchHaiti      =15372      # from enum MsoLanguageID
	msoLanguageIDFrenchLuxembourg =5132       # from enum MsoLanguageID
	msoLanguageIDFrenchMali       =13324      # from enum MsoLanguageID
	msoLanguageIDFrenchMonaco     =6156       # from enum MsoLanguageID
	msoLanguageIDFrenchMorocco    =14348      # from enum MsoLanguageID
	msoLanguageIDFrenchReunion    =8204       # from enum MsoLanguageID
	msoLanguageIDFrenchSenegal    =10252      # from enum MsoLanguageID
	msoLanguageIDFrenchWestIndies =7180       # from enum MsoLanguageID
	msoLanguageIDFrenchZaire      =9228       # from enum MsoLanguageID
	msoLanguageIDFrisianNetherlands=1122       # from enum MsoLanguageID
	msoLanguageIDFulfulde         =1127       # from enum MsoLanguageID
	msoLanguageIDGaelicIreland    =2108       # from enum MsoLanguageID
	msoLanguageIDGaelicScotland   =1084       # from enum MsoLanguageID
	msoLanguageIDGalician         =1110       # from enum MsoLanguageID
	msoLanguageIDGeorgian         =1079       # from enum MsoLanguageID
	msoLanguageIDGerman           =1031       # from enum MsoLanguageID
	msoLanguageIDGermanAustria    =3079       # from enum MsoLanguageID
	msoLanguageIDGermanLiechtenstein=5127       # from enum MsoLanguageID
	msoLanguageIDGermanLuxembourg =4103       # from enum MsoLanguageID
	msoLanguageIDGreek            =1032       # from enum MsoLanguageID
	msoLanguageIDGuarani          =1140       # from enum MsoLanguageID
	msoLanguageIDGujarati         =1095       # from enum MsoLanguageID
	msoLanguageIDHausa            =1128       # from enum MsoLanguageID
	msoLanguageIDHawaiian         =1141       # from enum MsoLanguageID
	msoLanguageIDHebrew           =1037       # from enum MsoLanguageID
	msoLanguageIDHindi            =1081       # from enum MsoLanguageID
	msoLanguageIDHungarian        =1038       # from enum MsoLanguageID
	msoLanguageIDIbibio           =1129       # from enum MsoLanguageID
	msoLanguageIDIcelandic        =1039       # from enum MsoLanguageID
	msoLanguageIDIgbo             =1136       # from enum MsoLanguageID
	msoLanguageIDIndonesian       =1057       # from enum MsoLanguageID
	msoLanguageIDInuktitut        =1117       # from enum MsoLanguageID
	msoLanguageIDItalian          =1040       # from enum MsoLanguageID
	msoLanguageIDJapanese         =1041       # from enum MsoLanguageID
	msoLanguageIDKannada          =1099       # from enum MsoLanguageID
	msoLanguageIDKanuri           =1137       # from enum MsoLanguageID
	msoLanguageIDKashmiri         =1120       # from enum MsoLanguageID
	msoLanguageIDKashmiriDevanagari=2144       # from enum MsoLanguageID
	msoLanguageIDKazakh           =1087       # from enum MsoLanguageID
	msoLanguageIDKhmer            =1107       # from enum MsoLanguageID
	msoLanguageIDKirghiz          =1088       # from enum MsoLanguageID
	msoLanguageIDKonkani          =1111       # from enum MsoLanguageID
	msoLanguageIDKorean           =1042       # from enum MsoLanguageID
	msoLanguageIDKyrgyz           =1088       # from enum MsoLanguageID
	msoLanguageIDLao              =1108       # from enum MsoLanguageID
	msoLanguageIDLatin            =1142       # from enum MsoLanguageID
	msoLanguageIDLatvian          =1062       # from enum MsoLanguageID
	msoLanguageIDLithuanian       =1063       # from enum MsoLanguageID
	msoLanguageIDMacedonian       =1071       # from enum MsoLanguageID
	msoLanguageIDMacedonianFYROM  =1071       # from enum MsoLanguageID
	msoLanguageIDMalayBruneiDarussalam=2110       # from enum MsoLanguageID
	msoLanguageIDMalayalam        =1100       # from enum MsoLanguageID
	msoLanguageIDMalaysian        =1086       # from enum MsoLanguageID
	msoLanguageIDMaltese          =1082       # from enum MsoLanguageID
	msoLanguageIDManipuri         =1112       # from enum MsoLanguageID
	msoLanguageIDMaori            =1153       # from enum MsoLanguageID
	msoLanguageIDMarathi          =1102       # from enum MsoLanguageID
	msoLanguageIDMexicanSpanish   =2058       # from enum MsoLanguageID
	msoLanguageIDMixed            =-2         # from enum MsoLanguageID
	msoLanguageIDMongolian        =1104       # from enum MsoLanguageID
	msoLanguageIDNepali           =1121       # from enum MsoLanguageID
	msoLanguageIDNoProofing       =1024       # from enum MsoLanguageID
	msoLanguageIDNone             =0          # from enum MsoLanguageID
	msoLanguageIDNorwegianBokmol  =1044       # from enum MsoLanguageID
	msoLanguageIDNorwegianNynorsk =2068       # from enum MsoLanguageID
	msoLanguageIDOriya            =1096       # from enum MsoLanguageID
	msoLanguageIDOromo            =1138       # from enum MsoLanguageID
	msoLanguageIDPashto           =1123       # from enum MsoLanguageID
	msoLanguageIDPolish           =1045       # from enum MsoLanguageID
	msoLanguageIDPortuguese       =2070       # from enum MsoLanguageID
	msoLanguageIDPunjabi          =1094       # from enum MsoLanguageID
	msoLanguageIDQuechuaBolivia   =1131       # from enum MsoLanguageID
	msoLanguageIDQuechuaEcuador   =2155       # from enum MsoLanguageID
	msoLanguageIDQuechuaPeru      =3179       # from enum MsoLanguageID
	msoLanguageIDRhaetoRomanic    =1047       # from enum MsoLanguageID
	msoLanguageIDRomanian         =1048       # from enum MsoLanguageID
	msoLanguageIDRomanianMoldova  =2072       # from enum MsoLanguageID
	msoLanguageIDRussian          =1049       # from enum MsoLanguageID
	msoLanguageIDRussianMoldova   =2073       # from enum MsoLanguageID
	msoLanguageIDSamiLappish      =1083       # from enum MsoLanguageID
	msoLanguageIDSanskrit         =1103       # from enum MsoLanguageID
	msoLanguageIDSepedi           =1132       # from enum MsoLanguageID
	msoLanguageIDSerbianBosniaHerzegovinaCyrillic=7194       # from enum MsoLanguageID
	msoLanguageIDSerbianBosniaHerzegovinaLatin=6170       # from enum MsoLanguageID
	msoLanguageIDSerbianCyrillic  =3098       # from enum MsoLanguageID
	msoLanguageIDSerbianLatin     =2074       # from enum MsoLanguageID
	msoLanguageIDSesotho          =1072       # from enum MsoLanguageID
	msoLanguageIDSimplifiedChinese=2052       # from enum MsoLanguageID
	msoLanguageIDSindhi           =1113       # from enum MsoLanguageID
	msoLanguageIDSindhiPakistan   =2137       # from enum MsoLanguageID
	msoLanguageIDSinhalese        =1115       # from enum MsoLanguageID
	msoLanguageIDSlovak           =1051       # from enum MsoLanguageID
	msoLanguageIDSlovenian        =1060       # from enum MsoLanguageID
	msoLanguageIDSomali           =1143       # from enum MsoLanguageID
	msoLanguageIDSorbian          =1070       # from enum MsoLanguageID
	msoLanguageIDSpanish          =1034       # from enum MsoLanguageID
	msoLanguageIDSpanishArgentina =11274      # from enum MsoLanguageID
	msoLanguageIDSpanishBolivia   =16394      # from enum MsoLanguageID
	msoLanguageIDSpanishChile     =13322      # from enum MsoLanguageID
	msoLanguageIDSpanishColombia  =9226       # from enum MsoLanguageID
	msoLanguageIDSpanishCostaRica =5130       # from enum MsoLanguageID
	msoLanguageIDSpanishDominicanRepublic=7178       # from enum MsoLanguageID
	msoLanguageIDSpanishEcuador   =12298      # from enum MsoLanguageID
	msoLanguageIDSpanishElSalvador=17418      # from enum MsoLanguageID
	msoLanguageIDSpanishGuatemala =4106       # from enum MsoLanguageID
	msoLanguageIDSpanishHonduras  =18442      # from enum MsoLanguageID
	msoLanguageIDSpanishModernSort=3082       # from enum MsoLanguageID
	msoLanguageIDSpanishNicaragua =19466      # from enum MsoLanguageID
	msoLanguageIDSpanishPanama    =6154       # from enum MsoLanguageID
	msoLanguageIDSpanishParaguay  =15370      # from enum MsoLanguageID
	msoLanguageIDSpanishPeru      =10250      # from enum MsoLanguageID
	msoLanguageIDSpanishPuertoRico=20490      # from enum MsoLanguageID
	msoLanguageIDSpanishUruguay   =14346      # from enum MsoLanguageID
	msoLanguageIDSpanishVenezuela =8202       # from enum MsoLanguageID
	msoLanguageIDSutu             =1072       # from enum MsoLanguageID
	msoLanguageIDSwahili          =1089       # from enum MsoLanguageID
	msoLanguageIDSwedish          =1053       # from enum MsoLanguageID
	msoLanguageIDSwedishFinland   =2077       # from enum MsoLanguageID
	msoLanguageIDSwissFrench      =4108       # from enum MsoLanguageID
	msoLanguageIDSwissGerman      =2055       # from enum MsoLanguageID
	msoLanguageIDSwissItalian     =2064       # from enum MsoLanguageID
	msoLanguageIDSyriac           =1114       # from enum MsoLanguageID
	msoLanguageIDTajik            =1064       # from enum MsoLanguageID
	msoLanguageIDTamazight        =1119       # from enum MsoLanguageID
	msoLanguageIDTamazightLatin   =2143       # from enum MsoLanguageID
	msoLanguageIDTamil            =1097       # from enum MsoLanguageID
	msoLanguageIDTatar            =1092       # from enum MsoLanguageID
	msoLanguageIDTelugu           =1098       # from enum MsoLanguageID
	msoLanguageIDThai             =1054       # from enum MsoLanguageID
	msoLanguageIDTibetan          =1105       # from enum MsoLanguageID
	msoLanguageIDTigrignaEritrea  =2163       # from enum MsoLanguageID
	msoLanguageIDTigrignaEthiopic =1139       # from enum MsoLanguageID
	msoLanguageIDTraditionalChinese=1028       # from enum MsoLanguageID
	msoLanguageIDTsonga           =1073       # from enum MsoLanguageID
	msoLanguageIDTswana           =1074       # from enum MsoLanguageID
	msoLanguageIDTurkish          =1055       # from enum MsoLanguageID
	msoLanguageIDTurkmen          =1090       # from enum MsoLanguageID
	msoLanguageIDUkrainian        =1058       # from enum MsoLanguageID
	msoLanguageIDUrdu             =1056       # from enum MsoLanguageID
	msoLanguageIDUzbekCyrillic    =2115       # from enum MsoLanguageID
	msoLanguageIDUzbekLatin       =1091       # from enum MsoLanguageID
	msoLanguageIDVenda            =1075       # from enum MsoLanguageID
	msoLanguageIDVietnamese       =1066       # from enum MsoLanguageID
	msoLanguageIDWelsh            =1106       # from enum MsoLanguageID
	msoLanguageIDXhosa            =1076       # from enum MsoLanguageID
	msoLanguageIDYi               =1144       # from enum MsoLanguageID
	msoLanguageIDYiddish          =1085       # from enum MsoLanguageID
	msoLanguageIDYoruba           =1130       # from enum MsoLanguageID
	msoLanguageIDZulu             =1077       # from enum MsoLanguageID
	msoLanguageIDChineseHongKong  =3076       # from enum MsoLanguageIDHidden
	msoLanguageIDChineseMacao     =5124       # from enum MsoLanguageIDHidden
	msoLanguageIDEnglishTrinidad  =11273      # from enum MsoLanguageIDHidden
	msoLastModifiedAnyTime        =7          # from enum MsoLastModified
	msoLastModifiedLastMonth      =5          # from enum MsoLastModified
	msoLastModifiedLastWeek       =3          # from enum MsoLastModified
	msoLastModifiedThisMonth      =6          # from enum MsoLastModified
	msoLastModifiedThisWeek       =4          # from enum MsoLastModified
	msoLastModifiedToday          =2          # from enum MsoLastModified
	msoLastModifiedYesterday      =1          # from enum MsoLastModified
	msoLightRigBalanced           =14         # from enum MsoLightRigType
	msoLightRigBrightRoom         =27         # from enum MsoLightRigType
	msoLightRigChilly             =22         # from enum MsoLightRigType
	msoLightRigContrasting        =18         # from enum MsoLightRigType
	msoLightRigFlat               =24         # from enum MsoLightRigType
	msoLightRigFlood              =17         # from enum MsoLightRigType
	msoLightRigFreezing           =23         # from enum MsoLightRigType
	msoLightRigGlow               =26         # from enum MsoLightRigType
	msoLightRigHarsh              =16         # from enum MsoLightRigType
	msoLightRigLegacyFlat1        =1          # from enum MsoLightRigType
	msoLightRigLegacyFlat2        =2          # from enum MsoLightRigType
	msoLightRigLegacyFlat3        =3          # from enum MsoLightRigType
	msoLightRigLegacyFlat4        =4          # from enum MsoLightRigType
	msoLightRigLegacyHarsh1       =9          # from enum MsoLightRigType
	msoLightRigLegacyHarsh2       =10         # from enum MsoLightRigType
	msoLightRigLegacyHarsh3       =11         # from enum MsoLightRigType
	msoLightRigLegacyHarsh4       =12         # from enum MsoLightRigType
	msoLightRigLegacyNormal1      =5          # from enum MsoLightRigType
	msoLightRigLegacyNormal2      =6          # from enum MsoLightRigType
	msoLightRigLegacyNormal3      =7          # from enum MsoLightRigType
	msoLightRigLegacyNormal4      =8          # from enum MsoLightRigType
	msoLightRigMixed              =-2         # from enum MsoLightRigType
	msoLightRigMorning            =19         # from enum MsoLightRigType
	msoLightRigSoft               =15         # from enum MsoLightRigType
	msoLightRigSunrise            =20         # from enum MsoLightRigType
	msoLightRigSunset             =21         # from enum MsoLightRigType
	msoLightRigThreePoint         =13         # from enum MsoLightRigType
	msoLightRigTwoPoint           =25         # from enum MsoLightRigType
	msoLineDash                   =4          # from enum MsoLineDashStyle
	msoLineDashDot                =5          # from enum MsoLineDashStyle
	msoLineDashDotDot             =6          # from enum MsoLineDashStyle
	msoLineDashStyleMixed         =-2         # from enum MsoLineDashStyle
	msoLineLongDash               =7          # from enum MsoLineDashStyle
	msoLineLongDashDot            =8          # from enum MsoLineDashStyle
	msoLineLongDashDotDot         =9          # from enum MsoLineDashStyle
	msoLineRoundDot               =3          # from enum MsoLineDashStyle
	msoLineSolid                  =1          # from enum MsoLineDashStyle
	msoLineSquareDot              =2          # from enum MsoLineDashStyle
	msoLineSysDash                =10         # from enum MsoLineDashStyle
	msoLineSysDashDot             =12         # from enum MsoLineDashStyle
	msoLineSysDot                 =11         # from enum MsoLineDashStyle
	msoLineSingle                 =1          # from enum MsoLineStyle
	msoLineStyleMixed             =-2         # from enum MsoLineStyle
	msoLineThickBetweenThin       =5          # from enum MsoLineStyle
	msoLineThickThin              =4          # from enum MsoLineStyle
	msoLineThinThick              =3          # from enum MsoLineStyle
	msoLineThinThin               =2          # from enum MsoLineStyle
	msoMenuAnimationNone          =0          # from enum MsoMenuAnimation
	msoMenuAnimationRandom        =1          # from enum MsoMenuAnimation
	msoMenuAnimationSlide         =3          # from enum MsoMenuAnimation
	msoMenuAnimationUnfold        =2          # from enum MsoMenuAnimation
	msoMetaPropertyTypeBoolean    =1          # from enum MsoMetaPropertyType
	msoMetaPropertyTypeBusinessData=20         # from enum MsoMetaPropertyType
	msoMetaPropertyTypeCalculated =3          # from enum MsoMetaPropertyType
	msoMetaPropertyTypeChoice     =2          # from enum MsoMetaPropertyType
	msoMetaPropertyTypeComputed   =4          # from enum MsoMetaPropertyType
	msoMetaPropertyTypeCurrency   =5          # from enum MsoMetaPropertyType
	msoMetaPropertyTypeDateTime   =6          # from enum MsoMetaPropertyType
	msoMetaPropertyTypeFillInChoice=7          # from enum MsoMetaPropertyType
	msoMetaPropertyTypeGuid       =8          # from enum MsoMetaPropertyType
	msoMetaPropertyTypeInteger    =9          # from enum MsoMetaPropertyType
	msoMetaPropertyTypeLookup     =10         # from enum MsoMetaPropertyType
	msoMetaPropertyTypeMax        =21         # from enum MsoMetaPropertyType
	msoMetaPropertyTypeMultiChoice=12         # from enum MsoMetaPropertyType
	msoMetaPropertyTypeMultiChoiceFillIn=13         # from enum MsoMetaPropertyType
	msoMetaPropertyTypeMultiChoiceLookup=11         # from enum MsoMetaPropertyType
	msoMetaPropertyTypeNote       =14         # from enum MsoMetaPropertyType
	msoMetaPropertyTypeNumber     =15         # from enum MsoMetaPropertyType
	msoMetaPropertyTypeText       =16         # from enum MsoMetaPropertyType
	msoMetaPropertyTypeUnknown    =0          # from enum MsoMetaPropertyType
	msoMetaPropertyTypeUrl        =17         # from enum MsoMetaPropertyType
	msoMetaPropertyTypeUser       =18         # from enum MsoMetaPropertyType
	msoMetaPropertyTypeUserMulti  =19         # from enum MsoMetaPropertyType
	msoIntegerMixed               =32768      # from enum MsoMixedType
	msoSingleMixed                =-2147483648 # from enum MsoMixedType
	msoModeAutoDown               =1          # from enum MsoModeType
	msoModeModal                  =0          # from enum MsoModeType
	msoModeModeless               =2          # from enum MsoModeType
	msoMoveRowFirst               =-4         # from enum MsoMoveRow
	msoMoveRowNbr                 =-1         # from enum MsoMoveRow
	msoMoveRowNext                =-2         # from enum MsoMoveRow
	msoMoveRowPrev                =-3         # from enum MsoMoveRow
	msoBulletAlphaLCParenBoth     =8          # from enum MsoNumberedBulletStyle
	msoBulletAlphaLCParenRight    =9          # from enum MsoNumberedBulletStyle
	msoBulletAlphaLCPeriod        =0          # from enum MsoNumberedBulletStyle
	msoBulletAlphaUCParenBoth     =10         # from enum MsoNumberedBulletStyle
	msoBulletAlphaUCParenRight    =11         # from enum MsoNumberedBulletStyle
	msoBulletAlphaUCPeriod        =1          # from enum MsoNumberedBulletStyle
	msoBulletArabicAbjadDash      =24         # from enum MsoNumberedBulletStyle
	msoBulletArabicAlphaDash      =23         # from enum MsoNumberedBulletStyle
	msoBulletArabicDBPeriod       =29         # from enum MsoNumberedBulletStyle
	msoBulletArabicDBPlain        =28         # from enum MsoNumberedBulletStyle
	msoBulletArabicParenBoth      =12         # from enum MsoNumberedBulletStyle
	msoBulletArabicParenRight     =2          # from enum MsoNumberedBulletStyle
	msoBulletArabicPeriod         =3          # from enum MsoNumberedBulletStyle
	msoBulletArabicPlain          =13         # from enum MsoNumberedBulletStyle
	msoBulletCircleNumDBPlain     =18         # from enum MsoNumberedBulletStyle
	msoBulletCircleNumWDBlackPlain=20         # from enum MsoNumberedBulletStyle
	msoBulletCircleNumWDWhitePlain=19         # from enum MsoNumberedBulletStyle
	msoBulletHebrewAlphaDash      =25         # from enum MsoNumberedBulletStyle
	msoBulletHindiAlpha1Period    =40         # from enum MsoNumberedBulletStyle
	msoBulletHindiAlphaPeriod     =36         # from enum MsoNumberedBulletStyle
	msoBulletHindiNumParenRight   =39         # from enum MsoNumberedBulletStyle
	msoBulletHindiNumPeriod       =37         # from enum MsoNumberedBulletStyle
	msoBulletKanjiKoreanPeriod    =27         # from enum MsoNumberedBulletStyle
	msoBulletKanjiKoreanPlain     =26         # from enum MsoNumberedBulletStyle
	msoBulletKanjiSimpChinDBPeriod=38         # from enum MsoNumberedBulletStyle
	msoBulletRomanLCParenBoth     =4          # from enum MsoNumberedBulletStyle
	msoBulletRomanLCParenRight    =5          # from enum MsoNumberedBulletStyle
	msoBulletRomanLCPeriod        =6          # from enum MsoNumberedBulletStyle
	msoBulletRomanUCParenBoth     =14         # from enum MsoNumberedBulletStyle
	msoBulletRomanUCParenRight    =15         # from enum MsoNumberedBulletStyle
	msoBulletRomanUCPeriod        =7          # from enum MsoNumberedBulletStyle
	msoBulletSimpChinPeriod       =17         # from enum MsoNumberedBulletStyle
	msoBulletSimpChinPlain        =16         # from enum MsoNumberedBulletStyle
	msoBulletStyleMixed           =-2         # from enum MsoNumberedBulletStyle
	msoBulletThaiAlphaParenBoth   =32         # from enum MsoNumberedBulletStyle
	msoBulletThaiAlphaParenRight  =31         # from enum MsoNumberedBulletStyle
	msoBulletThaiAlphaPeriod      =30         # from enum MsoNumberedBulletStyle
	msoBulletThaiNumParenBoth     =35         # from enum MsoNumberedBulletStyle
	msoBulletThaiNumParenRight    =34         # from enum MsoNumberedBulletStyle
	msoBulletThaiNumPeriod        =33         # from enum MsoNumberedBulletStyle
	msoBulletTradChinPeriod       =22         # from enum MsoNumberedBulletStyle
	msoBulletTradChinPlain        =21         # from enum MsoNumberedBulletStyle
	msoOLEMenuGroupContainer      =2          # from enum MsoOLEMenuGroup
	msoOLEMenuGroupEdit           =1          # from enum MsoOLEMenuGroup
	msoOLEMenuGroupFile           =0          # from enum MsoOLEMenuGroup
	msoOLEMenuGroupHelp           =5          # from enum MsoOLEMenuGroup
	msoOLEMenuGroupNone           =-1         # from enum MsoOLEMenuGroup
	msoOLEMenuGroupObject         =3          # from enum MsoOLEMenuGroup
	msoOLEMenuGroupWindow         =4          # from enum MsoOLEMenuGroup
	msoOrgChartLayoutBothHanging  =2          # from enum MsoOrgChartLayoutType
	msoOrgChartLayoutLeftHanging  =3          # from enum MsoOrgChartLayoutType
	msoOrgChartLayoutMixed        =-2         # from enum MsoOrgChartLayoutType
	msoOrgChartLayoutRightHanging =4          # from enum MsoOrgChartLayoutType
	msoOrgChartLayoutStandard     =1          # from enum MsoOrgChartLayoutType
	msoOrgChartOrientationMixed   =-2         # from enum MsoOrgChartOrientation
	msoOrgChartOrientationVertical=1          # from enum MsoOrgChartOrientation
	msoOrientationHorizontal      =1          # from enum MsoOrientation
	msoOrientationMixed           =-2         # from enum MsoOrientation
	msoOrientationVertical        =2          # from enum MsoOrientation
	msoAlignCenter                =2          # from enum MsoParagraphAlignment
	msoAlignDistribute            =5          # from enum MsoParagraphAlignment
	msoAlignJustify               =4          # from enum MsoParagraphAlignment
	msoAlignJustifyLow            =7          # from enum MsoParagraphAlignment
	msoAlignLeft                  =1          # from enum MsoParagraphAlignment
	msoAlignMixed                 =-2         # from enum MsoParagraphAlignment
	msoAlignRight                 =3          # from enum MsoParagraphAlignment
	msoAlignThaiDistribute        =6          # from enum MsoParagraphAlignment
	msoPathType1                  =1          # from enum MsoPathFormat
	msoPathType2                  =2          # from enum MsoPathFormat
	msoPathType3                  =3          # from enum MsoPathFormat
	msoPathType4                  =4          # from enum MsoPathFormat
	msoPathTypeMixed              =-2         # from enum MsoPathFormat
	msoPathTypeNone               =0          # from enum MsoPathFormat
	msoPattern10Percent           =2          # from enum MsoPatternType
	msoPattern20Percent           =3          # from enum MsoPatternType
	msoPattern25Percent           =4          # from enum MsoPatternType
	msoPattern30Percent           =5          # from enum MsoPatternType
	msoPattern40Percent           =6          # from enum MsoPatternType
	msoPattern50Percent           =7          # from enum MsoPatternType
	msoPattern5Percent            =1          # from enum MsoPatternType
	msoPattern60Percent           =8          # from enum MsoPatternType
	msoPattern70Percent           =9          # from enum MsoPatternType
	msoPattern75Percent           =10         # from enum MsoPatternType
	msoPattern80Percent           =11         # from enum MsoPatternType
	msoPattern90Percent           =12         # from enum MsoPatternType
	msoPatternCross               =51         # from enum MsoPatternType
	msoPatternDarkDownwardDiagonal=15         # from enum MsoPatternType
	msoPatternDarkHorizontal      =13         # from enum MsoPatternType
	msoPatternDarkUpwardDiagonal  =16         # from enum MsoPatternType
	msoPatternDarkVertical        =14         # from enum MsoPatternType
	msoPatternDashedDownwardDiagonal=28         # from enum MsoPatternType
	msoPatternDashedHorizontal    =32         # from enum MsoPatternType
	msoPatternDashedUpwardDiagonal=27         # from enum MsoPatternType
	msoPatternDashedVertical      =31         # from enum MsoPatternType
	msoPatternDiagonalBrick       =40         # from enum MsoPatternType
	msoPatternDiagonalCross       =54         # from enum MsoPatternType
	msoPatternDivot               =46         # from enum MsoPatternType
	msoPatternDottedDiamond       =24         # from enum MsoPatternType
	msoPatternDottedGrid          =45         # from enum MsoPatternType
	msoPatternDownwardDiagonal    =52         # from enum MsoPatternType
	msoPatternHorizontal          =49         # from enum MsoPatternType
	msoPatternHorizontalBrick     =35         # from enum MsoPatternType
	msoPatternLargeCheckerBoard   =36         # from enum MsoPatternType
	msoPatternLargeConfetti       =33         # from enum MsoPatternType
	msoPatternLargeGrid           =34         # from enum MsoPatternType
	msoPatternLightDownwardDiagonal=21         # from enum MsoPatternType
	msoPatternLightHorizontal     =19         # from enum MsoPatternType
	msoPatternLightUpwardDiagonal =22         # from enum MsoPatternType
	msoPatternLightVertical       =20         # from enum MsoPatternType
	msoPatternMixed               =-2         # from enum MsoPatternType
	msoPatternNarrowHorizontal    =30         # from enum MsoPatternType
	msoPatternNarrowVertical      =29         # from enum MsoPatternType
	msoPatternOutlinedDiamond     =41         # from enum MsoPatternType
	msoPatternPlaid               =42         # from enum MsoPatternType
	msoPatternShingle             =47         # from enum MsoPatternType
	msoPatternSmallCheckerBoard   =17         # from enum MsoPatternType
	msoPatternSmallConfetti       =37         # from enum MsoPatternType
	msoPatternSmallGrid           =23         # from enum MsoPatternType
	msoPatternSolidDiamond        =39         # from enum MsoPatternType
	msoPatternSphere              =43         # from enum MsoPatternType
	msoPatternTrellis             =18         # from enum MsoPatternType
	msoPatternUpwardDiagonal      =53         # from enum MsoPatternType
	msoPatternVertical            =50         # from enum MsoPatternType
	msoPatternWave                =48         # from enum MsoPatternType
	msoPatternWeave               =44         # from enum MsoPatternType
	msoPatternWideDownwardDiagonal=25         # from enum MsoPatternType
	msoPatternWideUpwardDiagonal  =26         # from enum MsoPatternType
	msoPatternZigZag              =38         # from enum MsoPatternType
	msoPermissionAllCommon        =127        # from enum MsoPermission
	msoPermissionChange           =15         # from enum MsoPermission
	msoPermissionEdit             =2          # from enum MsoPermission
	msoPermissionExtract          =8          # from enum MsoPermission
	msoPermissionFullControl      =64         # from enum MsoPermission
	msoPermissionObjModel         =32         # from enum MsoPermission
	msoPermissionPrint            =16         # from enum MsoPermission
	msoPermissionRead             =1          # from enum MsoPermission
	msoPermissionSave             =4          # from enum MsoPermission
	msoPermissionView             =1          # from enum MsoPermission
	msoPictureAutomatic           =1          # from enum MsoPictureColorType
	msoPictureBlackAndWhite       =3          # from enum MsoPictureColorType
	msoPictureGrayscale           =2          # from enum MsoPictureColorType
	msoPictureMixed               =-2         # from enum MsoPictureColorType
	msoPictureWatermark           =4          # from enum MsoPictureColorType
	msoCameraIsometricBottomDown  =23         # from enum MsoPresetCamera
	msoCameraIsometricBottomUp    =22         # from enum MsoPresetCamera
	msoCameraIsometricLeftDown    =25         # from enum MsoPresetCamera
	msoCameraIsometricLeftUp      =24         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis1Left=28         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis1Right=29         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis1Top =30         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis2Left=31         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis2Right=32         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis2Top =33         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis3Bottom=36         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis3Left=34         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis3Right=35         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis4Bottom=39         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis4Left=37         # from enum MsoPresetCamera
	msoCameraIsometricOffAxis4Right=38         # from enum MsoPresetCamera
	msoCameraIsometricRightDown   =27         # from enum MsoPresetCamera
	msoCameraIsometricRightUp     =26         # from enum MsoPresetCamera
	msoCameraIsometricTopDown     =21         # from enum MsoPresetCamera
	msoCameraIsometricTopUp       =20         # from enum MsoPresetCamera
	msoCameraLegacyObliqueBottom  =8          # from enum MsoPresetCamera
	msoCameraLegacyObliqueBottomLeft=7          # from enum MsoPresetCamera
	msoCameraLegacyObliqueBottomRight=9          # from enum MsoPresetCamera
	msoCameraLegacyObliqueFront   =5          # from enum MsoPresetCamera
	msoCameraLegacyObliqueLeft    =4          # from enum MsoPresetCamera
	msoCameraLegacyObliqueRight   =6          # from enum MsoPresetCamera
	msoCameraLegacyObliqueTop     =2          # from enum MsoPresetCamera
	msoCameraLegacyObliqueTopLeft =1          # from enum MsoPresetCamera
	msoCameraLegacyObliqueTopRight=3          # from enum MsoPresetCamera
	msoCameraLegacyPerspectiveBottom=17         # from enum MsoPresetCamera
	msoCameraLegacyPerspectiveBottomLeft=16         # from enum MsoPresetCamera
	msoCameraLegacyPerspectiveBottomRight=18         # from enum MsoPresetCamera
	msoCameraLegacyPerspectiveFront=14         # from enum MsoPresetCamera
	msoCameraLegacyPerspectiveLeft=13         # from enum MsoPresetCamera
	msoCameraLegacyPerspectiveRight=15         # from enum MsoPresetCamera
	msoCameraLegacyPerspectiveTop =11         # from enum MsoPresetCamera
	msoCameraLegacyPerspectiveTopLeft=10         # from enum MsoPresetCamera
	msoCameraLegacyPerspectiveTopRight=12         # from enum MsoPresetCamera
	msoCameraObliqueBottom        =46         # from enum MsoPresetCamera
	msoCameraObliqueBottomLeft    =45         # from enum MsoPresetCamera
	msoCameraObliqueBottomRight   =47         # from enum MsoPresetCamera
	msoCameraObliqueLeft          =43         # from enum MsoPresetCamera
	msoCameraObliqueRight         =44         # from enum MsoPresetCamera
	msoCameraObliqueTop           =41         # from enum MsoPresetCamera
	msoCameraObliqueTopLeft       =40         # from enum MsoPresetCamera
	msoCameraObliqueTopRight      =42         # from enum MsoPresetCamera
	msoCameraOrthographicFront    =19         # from enum MsoPresetCamera
	msoCameraPerspectiveAbove     =51         # from enum MsoPresetCamera
	msoCameraPerspectiveAboveLeftFacing=53         # from enum MsoPresetCamera
	msoCameraPerspectiveAboveRightFacing=54         # from enum MsoPresetCamera
	msoCameraPerspectiveBelow     =52         # from enum MsoPresetCamera
	msoCameraPerspectiveContrastingLeftFacing=55         # from enum MsoPresetCamera
	msoCameraPerspectiveContrastingRightFacing=56         # from enum MsoPresetCamera
	msoCameraPerspectiveFront     =48         # from enum MsoPresetCamera
	msoCameraPerspectiveHeroicExtremeLeftFacing=59         # from enum MsoPresetCamera
	msoCameraPerspectiveHeroicExtremeRightFacing=60         # from enum MsoPresetCamera
	msoCameraPerspectiveHeroicLeftFacing=57         # from enum MsoPresetCamera
	msoCameraPerspectiveHeroicRightFacing=58         # from enum MsoPresetCamera
	msoCameraPerspectiveLeft      =49         # from enum MsoPresetCamera
	msoCameraPerspectiveRelaxed   =61         # from enum MsoPresetCamera
	msoCameraPerspectiveRelaxedModerately=62         # from enum MsoPresetCamera
	msoCameraPerspectiveRight     =50         # from enum MsoPresetCamera
	msoPresetCameraMixed          =-2         # from enum MsoPresetCamera
	msoExtrusionBottom            =2          # from enum MsoPresetExtrusionDirection
	msoExtrusionBottomLeft        =3          # from enum MsoPresetExtrusionDirection
	msoExtrusionBottomRight       =1          # from enum MsoPresetExtrusionDirection
	msoExtrusionLeft              =6          # from enum MsoPresetExtrusionDirection
	msoExtrusionNone              =5          # from enum MsoPresetExtrusionDirection
	msoExtrusionRight             =4          # from enum MsoPresetExtrusionDirection
	msoExtrusionTop               =8          # from enum MsoPresetExtrusionDirection
	msoExtrusionTopLeft           =9          # from enum MsoPresetExtrusionDirection
	msoExtrusionTopRight          =7          # from enum MsoPresetExtrusionDirection
	msoPresetExtrusionDirectionMixed=-2         # from enum MsoPresetExtrusionDirection
	msoGradientBrass              =20         # from enum MsoPresetGradientType
	msoGradientCalmWater          =8          # from enum MsoPresetGradientType
	msoGradientChrome             =21         # from enum MsoPresetGradientType
	msoGradientChromeII           =22         # from enum MsoPresetGradientType
	msoGradientDaybreak           =4          # from enum MsoPresetGradientType
	msoGradientDesert             =6          # from enum MsoPresetGradientType
	msoGradientEarlySunset        =1          # from enum MsoPresetGradientType
	msoGradientFire               =9          # from enum MsoPresetGradientType
	msoGradientFog                =10         # from enum MsoPresetGradientType
	msoGradientGold               =18         # from enum MsoPresetGradientType
	msoGradientGoldII             =19         # from enum MsoPresetGradientType
	msoGradientHorizon            =5          # from enum MsoPresetGradientType
	msoGradientLateSunset         =2          # from enum MsoPresetGradientType
	msoGradientMahogany           =15         # from enum MsoPresetGradientType
	msoGradientMoss               =11         # from enum MsoPresetGradientType
	msoGradientNightfall          =3          # from enum MsoPresetGradientType
	msoGradientOcean              =7          # from enum MsoPresetGradientType
	msoGradientParchment          =14         # from enum MsoPresetGradientType
	msoGradientPeacock            =12         # from enum MsoPresetGradientType
	msoGradientRainbow            =16         # from enum MsoPresetGradientType
	msoGradientRainbowII          =17         # from enum MsoPresetGradientType
	msoGradientSapphire           =24         # from enum MsoPresetGradientType
	msoGradientSilver             =23         # from enum MsoPresetGradientType
	msoGradientWheat              =13         # from enum MsoPresetGradientType
	msoPresetGradientMixed        =-2         # from enum MsoPresetGradientType
	msoLightingBottom             =8          # from enum MsoPresetLightingDirection
	msoLightingBottomLeft         =7          # from enum MsoPresetLightingDirection
	msoLightingBottomRight        =9          # from enum MsoPresetLightingDirection
	msoLightingLeft               =4          # from enum MsoPresetLightingDirection
	msoLightingNone               =5          # from enum MsoPresetLightingDirection
	msoLightingRight              =6          # from enum MsoPresetLightingDirection
	msoLightingTop                =2          # from enum MsoPresetLightingDirection
	msoLightingTopLeft            =1          # from enum MsoPresetLightingDirection
	msoLightingTopRight           =3          # from enum MsoPresetLightingDirection
	msoPresetLightingDirectionMixed=-2         # from enum MsoPresetLightingDirection
	msoLightingBright             =3          # from enum MsoPresetLightingSoftness
	msoLightingDim                =1          # from enum MsoPresetLightingSoftness
	msoLightingNormal             =2          # from enum MsoPresetLightingSoftness
	msoPresetLightingSoftnessMixed=-2         # from enum MsoPresetLightingSoftness
	msoMaterialClear              =13         # from enum MsoPresetMaterial
	msoMaterialDarkEdge           =11         # from enum MsoPresetMaterial
	msoMaterialFlat               =14         # from enum MsoPresetMaterial
	msoMaterialMatte              =1          # from enum MsoPresetMaterial
	msoMaterialMatte2             =5          # from enum MsoPresetMaterial
	msoMaterialMetal              =3          # from enum MsoPresetMaterial
	msoMaterialMetal2             =7          # from enum MsoPresetMaterial
	msoMaterialPlastic            =2          # from enum MsoPresetMaterial
	msoMaterialPlastic2           =6          # from enum MsoPresetMaterial
	msoMaterialPowder             =10         # from enum MsoPresetMaterial
	msoMaterialSoftEdge           =12         # from enum MsoPresetMaterial
	msoMaterialSoftMetal          =15         # from enum MsoPresetMaterial
	msoMaterialTranslucentPowder  =9          # from enum MsoPresetMaterial
	msoMaterialWarmMatte          =8          # from enum MsoPresetMaterial
	msoMaterialWireFrame          =4          # from enum MsoPresetMaterial
	msoPresetMaterialMixed        =-2         # from enum MsoPresetMaterial
	msoTextEffect1                =0          # from enum MsoPresetTextEffect
	msoTextEffect10               =9          # from enum MsoPresetTextEffect
	msoTextEffect11               =10         # from enum MsoPresetTextEffect
	msoTextEffect12               =11         # from enum MsoPresetTextEffect
	msoTextEffect13               =12         # from enum MsoPresetTextEffect
	msoTextEffect14               =13         # from enum MsoPresetTextEffect
	msoTextEffect15               =14         # from enum MsoPresetTextEffect
	msoTextEffect16               =15         # from enum MsoPresetTextEffect
	msoTextEffect17               =16         # from enum MsoPresetTextEffect
	msoTextEffect18               =17         # from enum MsoPresetTextEffect
	msoTextEffect19               =18         # from enum MsoPresetTextEffect
	msoTextEffect2                =1          # from enum MsoPresetTextEffect
	msoTextEffect20               =19         # from enum MsoPresetTextEffect
	msoTextEffect21               =20         # from enum MsoPresetTextEffect
	msoTextEffect22               =21         # from enum MsoPresetTextEffect
	msoTextEffect23               =22         # from enum MsoPresetTextEffect
	msoTextEffect24               =23         # from enum MsoPresetTextEffect
	msoTextEffect25               =24         # from enum MsoPresetTextEffect
	msoTextEffect26               =25         # from enum MsoPresetTextEffect
	msoTextEffect27               =26         # from enum MsoPresetTextEffect
	msoTextEffect28               =27         # from enum MsoPresetTextEffect
	msoTextEffect29               =28         # from enum MsoPresetTextEffect
	msoTextEffect3                =2          # from enum MsoPresetTextEffect
	msoTextEffect30               =29         # from enum MsoPresetTextEffect
	msoTextEffect4                =3          # from enum MsoPresetTextEffect
	msoTextEffect5                =4          # from enum MsoPresetTextEffect
	msoTextEffect6                =5          # from enum MsoPresetTextEffect
	msoTextEffect7                =6          # from enum MsoPresetTextEffect
	msoTextEffect8                =7          # from enum MsoPresetTextEffect
	msoTextEffect9                =8          # from enum MsoPresetTextEffect
	msoTextEffectMixed            =-2         # from enum MsoPresetTextEffect
	msoTextEffectShapeArchDownCurve=10         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeArchDownPour=14         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeArchUpCurve =9          # from enum MsoPresetTextEffectShape
	msoTextEffectShapeArchUpPour  =13         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeButtonCurve =12         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeButtonPour  =16         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeCanDown     =20         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeCanUp       =19         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeCascadeDown =40         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeCascadeUp   =39         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeChevronDown =6          # from enum MsoPresetTextEffectShape
	msoTextEffectShapeChevronUp   =5          # from enum MsoPresetTextEffectShape
	msoTextEffectShapeCircleCurve =11         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeCirclePour  =15         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeCurveDown   =18         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeCurveUp     =17         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeDeflate     =26         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeDeflateBottom=28         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeDeflateInflate=31         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeDeflateInflateDeflate=32         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeDeflateTop  =30         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeDoubleWave1 =23         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeDoubleWave2 =24         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeFadeDown    =36         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeFadeLeft    =34         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeFadeRight   =33         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeFadeUp      =35         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeInflate     =25         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeInflateBottom=27         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeInflateTop  =29         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeMixed       =-2         # from enum MsoPresetTextEffectShape
	msoTextEffectShapePlainText   =1          # from enum MsoPresetTextEffectShape
	msoTextEffectShapeRingInside  =7          # from enum MsoPresetTextEffectShape
	msoTextEffectShapeRingOutside =8          # from enum MsoPresetTextEffectShape
	msoTextEffectShapeSlantDown   =38         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeSlantUp     =37         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeStop        =2          # from enum MsoPresetTextEffectShape
	msoTextEffectShapeTriangleDown=4          # from enum MsoPresetTextEffectShape
	msoTextEffectShapeTriangleUp  =3          # from enum MsoPresetTextEffectShape
	msoTextEffectShapeWave1       =21         # from enum MsoPresetTextEffectShape
	msoTextEffectShapeWave2       =22         # from enum MsoPresetTextEffectShape
	msoPresetTextureMixed         =-2         # from enum MsoPresetTexture
	msoTextureBlueTissuePaper     =17         # from enum MsoPresetTexture
	msoTextureBouquet             =20         # from enum MsoPresetTexture
	msoTextureBrownMarble         =11         # from enum MsoPresetTexture
	msoTextureCanvas              =2          # from enum MsoPresetTexture
	msoTextureCork                =21         # from enum MsoPresetTexture
	msoTextureDenim               =3          # from enum MsoPresetTexture
	msoTextureFishFossil          =7          # from enum MsoPresetTexture
	msoTextureGranite             =12         # from enum MsoPresetTexture
	msoTextureGreenMarble         =9          # from enum MsoPresetTexture
	msoTextureMediumWood          =24         # from enum MsoPresetTexture
	msoTextureNewsprint           =13         # from enum MsoPresetTexture
	msoTextureOak                 =23         # from enum MsoPresetTexture
	msoTexturePaperBag            =6          # from enum MsoPresetTexture
	msoTexturePapyrus             =1          # from enum MsoPresetTexture
	msoTextureParchment           =15         # from enum MsoPresetTexture
	msoTexturePinkTissuePaper     =18         # from enum MsoPresetTexture
	msoTexturePurpleMesh          =19         # from enum MsoPresetTexture
	msoTextureRecycledPaper       =14         # from enum MsoPresetTexture
	msoTextureSand                =8          # from enum MsoPresetTexture
	msoTextureStationery          =16         # from enum MsoPresetTexture
	msoTextureWalnut              =22         # from enum MsoPresetTexture
	msoTextureWaterDroplets       =5          # from enum MsoPresetTexture
	msoTextureWhiteMarble         =10         # from enum MsoPresetTexture
	msoTextureWovenMat            =4          # from enum MsoPresetTexture
	msoPresetThreeDFormatMixed    =-2         # from enum MsoPresetThreeDFormat
	msoThreeD1                    =1          # from enum MsoPresetThreeDFormat
	msoThreeD10                   =10         # from enum MsoPresetThreeDFormat
	msoThreeD11                   =11         # from enum MsoPresetThreeDFormat
	msoThreeD12                   =12         # from enum MsoPresetThreeDFormat
	msoThreeD13                   =13         # from enum MsoPresetThreeDFormat
	msoThreeD14                   =14         # from enum MsoPresetThreeDFormat
	msoThreeD15                   =15         # from enum MsoPresetThreeDFormat
	msoThreeD16                   =16         # from enum MsoPresetThreeDFormat
	msoThreeD17                   =17         # from enum MsoPresetThreeDFormat
	msoThreeD18                   =18         # from enum MsoPresetThreeDFormat
	msoThreeD19                   =19         # from enum MsoPresetThreeDFormat
	msoThreeD2                    =2          # from enum MsoPresetThreeDFormat
	msoThreeD20                   =20         # from enum MsoPresetThreeDFormat
	msoThreeD3                    =3          # from enum MsoPresetThreeDFormat
	msoThreeD4                    =4          # from enum MsoPresetThreeDFormat
	msoThreeD5                    =5          # from enum MsoPresetThreeDFormat
	msoThreeD6                    =6          # from enum MsoPresetThreeDFormat
	msoThreeD7                    =7          # from enum MsoPresetThreeDFormat
	msoThreeD8                    =8          # from enum MsoPresetThreeDFormat
	msoThreeD9                    =9          # from enum MsoPresetThreeDFormat
	msoReflectionType1            =1          # from enum MsoReflectionType
	msoReflectionType2            =2          # from enum MsoReflectionType
	msoReflectionType3            =3          # from enum MsoReflectionType
	msoReflectionType4            =4          # from enum MsoReflectionType
	msoReflectionType5            =5          # from enum MsoReflectionType
	msoReflectionType6            =6          # from enum MsoReflectionType
	msoReflectionType7            =7          # from enum MsoReflectionType
	msoReflectionType8            =8          # from enum MsoReflectionType
	msoReflectionType9            =9          # from enum MsoReflectionType
	msoReflectionTypeMixed        =-2         # from enum MsoReflectionType
	msoReflectionTypeNone         =0          # from enum MsoReflectionType
	msoAfterLastSibling           =4          # from enum MsoRelativeNodePosition
	msoAfterNode                  =2          # from enum MsoRelativeNodePosition
	msoBeforeFirstSibling         =3          # from enum MsoRelativeNodePosition
	msoBeforeNode                 =1          # from enum MsoRelativeNodePosition
	msoScaleFromBottomRight       =2          # from enum MsoScaleFrom
	msoScaleFromMiddle            =1          # from enum MsoScaleFrom
	msoScaleFromTopLeft           =0          # from enum MsoScaleFrom
	msoScreenSize1024x768         =4          # from enum MsoScreenSize
	msoScreenSize1152x882         =5          # from enum MsoScreenSize
	msoScreenSize1152x900         =6          # from enum MsoScreenSize
	msoScreenSize1280x1024        =7          # from enum MsoScreenSize
	msoScreenSize1600x1200        =8          # from enum MsoScreenSize
	msoScreenSize1800x1440        =9          # from enum MsoScreenSize
	msoScreenSize1920x1200        =10         # from enum MsoScreenSize
	msoScreenSize544x376          =0          # from enum MsoScreenSize
	msoScreenSize640x480          =1          # from enum MsoScreenSize
	msoScreenSize720x512          =2          # from enum MsoScreenSize
	msoScreenSize800x600          =3          # from enum MsoScreenSize
	msoScriptLanguageASP          =3          # from enum MsoScriptLanguage
	msoScriptLanguageJava         =1          # from enum MsoScriptLanguage
	msoScriptLanguageOther        =4          # from enum MsoScriptLanguage
	msoScriptLanguageVisualBasic  =2          # from enum MsoScriptLanguage
	msoScriptLocationInBody       =2          # from enum MsoScriptLocation
	msoScriptLocationInHead       =1          # from enum MsoScriptLocation
	msoSearchInCustom             =3          # from enum MsoSearchIn
	msoSearchInMyComputer         =0          # from enum MsoSearchIn
	msoSearchInMyNetworkPlaces    =2          # from enum MsoSearchIn
	msoSearchInOutlook            =1          # from enum MsoSearchIn
	msoSegmentCurve               =1          # from enum MsoSegmentType
	msoSegmentLine                =0          # from enum MsoSegmentType
	msoShadowStyleInnerShadow     =1          # from enum MsoShadowStyle
	msoShadowStyleMixed           =-2         # from enum MsoShadowStyle
	msoShadowStyleOuterShadow     =2          # from enum MsoShadowStyle
	msoShadow1                    =1          # from enum MsoShadowType
	msoShadow10                   =10         # from enum MsoShadowType
	msoShadow11                   =11         # from enum MsoShadowType
	msoShadow12                   =12         # from enum MsoShadowType
	msoShadow13                   =13         # from enum MsoShadowType
	msoShadow14                   =14         # from enum MsoShadowType
	msoShadow15                   =15         # from enum MsoShadowType
	msoShadow16                   =16         # from enum MsoShadowType
	msoShadow17                   =17         # from enum MsoShadowType
	msoShadow18                   =18         # from enum MsoShadowType
	msoShadow19                   =19         # from enum MsoShadowType
	msoShadow2                    =2          # from enum MsoShadowType
	msoShadow20                   =20         # from enum MsoShadowType
	msoShadow3                    =3          # from enum MsoShadowType
	msoShadow4                    =4          # from enum MsoShadowType
	msoShadow5                    =5          # from enum MsoShadowType
	msoShadow6                    =6          # from enum MsoShadowType
	msoShadow7                    =7          # from enum MsoShadowType
	msoShadow8                    =8          # from enum MsoShadowType
	msoShadow9                    =9          # from enum MsoShadowType
	msoShadowMixed                =-2         # from enum MsoShadowType
	msoLineStylePreset1           =10001      # from enum MsoShapeStyleIndex
	msoLineStylePreset10          =10010      # from enum MsoShapeStyleIndex
	msoLineStylePreset11          =10011      # from enum MsoShapeStyleIndex
	msoLineStylePreset12          =10012      # from enum MsoShapeStyleIndex
	msoLineStylePreset13          =10013      # from enum MsoShapeStyleIndex
	msoLineStylePreset14          =10014      # from enum MsoShapeStyleIndex
	msoLineStylePreset15          =10015      # from enum MsoShapeStyleIndex
	msoLineStylePreset16          =10016      # from enum MsoShapeStyleIndex
	msoLineStylePreset17          =10017      # from enum MsoShapeStyleIndex
	msoLineStylePreset18          =10018      # from enum MsoShapeStyleIndex
	msoLineStylePreset19          =10019      # from enum MsoShapeStyleIndex
	msoLineStylePreset2           =10002      # from enum MsoShapeStyleIndex
	msoLineStylePreset20          =10020      # from enum MsoShapeStyleIndex
	msoLineStylePreset21          =10021      # from enum MsoShapeStyleIndex
	msoLineStylePreset3           =10003      # from enum MsoShapeStyleIndex
	msoLineStylePreset4           =10004      # from enum MsoShapeStyleIndex
	msoLineStylePreset5           =10005      # from enum MsoShapeStyleIndex
	msoLineStylePreset6           =10006      # from enum MsoShapeStyleIndex
	msoLineStylePreset7           =10007      # from enum MsoShapeStyleIndex
	msoLineStylePreset8           =10008      # from enum MsoShapeStyleIndex
	msoLineStylePreset9           =10009      # from enum MsoShapeStyleIndex
	msoShapeStyleMixed            =-2         # from enum MsoShapeStyleIndex
	msoShapeStyleNotAPreset       =0          # from enum MsoShapeStyleIndex
	msoShapeStylePreset1          =1          # from enum MsoShapeStyleIndex
	msoShapeStylePreset10         =10         # from enum MsoShapeStyleIndex
	msoShapeStylePreset11         =11         # from enum MsoShapeStyleIndex
	msoShapeStylePreset12         =12         # from enum MsoShapeStyleIndex
	msoShapeStylePreset13         =13         # from enum MsoShapeStyleIndex
	msoShapeStylePreset14         =14         # from enum MsoShapeStyleIndex
	msoShapeStylePreset15         =15         # from enum MsoShapeStyleIndex
	msoShapeStylePreset16         =16         # from enum MsoShapeStyleIndex
	msoShapeStylePreset17         =17         # from enum MsoShapeStyleIndex
	msoShapeStylePreset18         =18         # from enum MsoShapeStyleIndex
	msoShapeStylePreset19         =19         # from enum MsoShapeStyleIndex
	msoShapeStylePreset2          =2          # from enum MsoShapeStyleIndex
	msoShapeStylePreset20         =20         # from enum MsoShapeStyleIndex
	msoShapeStylePreset21         =21         # from enum MsoShapeStyleIndex
	msoShapeStylePreset22         =22         # from enum MsoShapeStyleIndex
	msoShapeStylePreset23         =23         # from enum MsoShapeStyleIndex
	msoShapeStylePreset24         =24         # from enum MsoShapeStyleIndex
	msoShapeStylePreset25         =25         # from enum MsoShapeStyleIndex
	msoShapeStylePreset26         =26         # from enum MsoShapeStyleIndex
	msoShapeStylePreset27         =27         # from enum MsoShapeStyleIndex
	msoShapeStylePreset28         =28         # from enum MsoShapeStyleIndex
	msoShapeStylePreset29         =29         # from enum MsoShapeStyleIndex
	msoShapeStylePreset3          =3          # from enum MsoShapeStyleIndex
	msoShapeStylePreset30         =30         # from enum MsoShapeStyleIndex
	msoShapeStylePreset31         =31         # from enum MsoShapeStyleIndex
	msoShapeStylePreset32         =32         # from enum MsoShapeStyleIndex
	msoShapeStylePreset33         =33         # from enum MsoShapeStyleIndex
	msoShapeStylePreset34         =34         # from enum MsoShapeStyleIndex
	msoShapeStylePreset35         =35         # from enum MsoShapeStyleIndex
	msoShapeStylePreset36         =36         # from enum MsoShapeStyleIndex
	msoShapeStylePreset37         =37         # from enum MsoShapeStyleIndex
	msoShapeStylePreset38         =38         # from enum MsoShapeStyleIndex
	msoShapeStylePreset39         =39         # from enum MsoShapeStyleIndex
	msoShapeStylePreset4          =4          # from enum MsoShapeStyleIndex
	msoShapeStylePreset40         =40         # from enum MsoShapeStyleIndex
	msoShapeStylePreset41         =41         # from enum MsoShapeStyleIndex
	msoShapeStylePreset42         =42         # from enum MsoShapeStyleIndex
	msoShapeStylePreset5          =5          # from enum MsoShapeStyleIndex
	msoShapeStylePreset6          =6          # from enum MsoShapeStyleIndex
	msoShapeStylePreset7          =7          # from enum MsoShapeStyleIndex
	msoShapeStylePreset8          =8          # from enum MsoShapeStyleIndex
	msoShapeStylePreset9          =9          # from enum MsoShapeStyleIndex
	msoAutoShape                  =1          # from enum MsoShapeType
	msoCallout                    =2          # from enum MsoShapeType
	msoCanvas                     =20         # from enum MsoShapeType
	msoChart                      =3          # from enum MsoShapeType
	msoComment                    =4          # from enum MsoShapeType
	msoDiagram                    =21         # from enum MsoShapeType
	msoEmbeddedOLEObject          =7          # from enum MsoShapeType
	msoFormControl                =8          # from enum MsoShapeType
	msoFreeform                   =5          # from enum MsoShapeType
	msoGroup                      =6          # from enum MsoShapeType
	msoInk                        =22         # from enum MsoShapeType
	msoInkComment                 =23         # from enum MsoShapeType
	msoLine                       =9          # from enum MsoShapeType
	msoLinkedOLEObject            =10         # from enum MsoShapeType
	msoLinkedPicture              =11         # from enum MsoShapeType
	msoMedia                      =16         # from enum MsoShapeType
	msoOLEControlObject           =12         # from enum MsoShapeType
	msoPicture                    =13         # from enum MsoShapeType
	msoPlaceholder                =14         # from enum MsoShapeType
	msoScriptAnchor               =18         # from enum MsoShapeType
	msoShapeTypeMixed             =-2         # from enum MsoShapeType
	msoSmartArt                   =24         # from enum MsoShapeType
	msoTable                      =19         # from enum MsoShapeType
	msoTextBox                    =17         # from enum MsoShapeType
	msoTextEffect                 =15         # from enum MsoShapeType
	msoSharedWorkspaceTaskPriorityHigh=1          # from enum MsoSharedWorkspaceTaskPriority
	msoSharedWorkspaceTaskPriorityLow=3          # from enum MsoSharedWorkspaceTaskPriority
	msoSharedWorkspaceTaskPriorityNormal=2          # from enum MsoSharedWorkspaceTaskPriority
	msoSharedWorkspaceTaskStatusCompleted=3          # from enum MsoSharedWorkspaceTaskStatus
	msoSharedWorkspaceTaskStatusDeferred=4          # from enum MsoSharedWorkspaceTaskStatus
	msoSharedWorkspaceTaskStatusInProgress=2          # from enum MsoSharedWorkspaceTaskStatus
	msoSharedWorkspaceTaskStatusNotStarted=1          # from enum MsoSharedWorkspaceTaskStatus
	msoSharedWorkspaceTaskStatusWaiting=5          # from enum MsoSharedWorkspaceTaskStatus
	msoSignatureSubsetAll         =5          # from enum MsoSignatureSubset
	msoSignatureSubsetSignatureLines=2          # from enum MsoSignatureSubset
	msoSignatureSubsetSignatureLinesSigned=3          # from enum MsoSignatureSubset
	msoSignatureSubsetSignatureLinesUnsigned=4          # from enum MsoSignatureSubset
	msoSignatureSubsetSignaturesAllSigs=0          # from enum MsoSignatureSubset
	msoSignatureSubsetSignaturesNonVisible=1          # from enum MsoSignatureSubset
	msoSoftEdgeType1              =1          # from enum MsoSoftEdgeType
	msoSoftEdgeType2              =2          # from enum MsoSoftEdgeType
	msoSoftEdgeType3              =3          # from enum MsoSoftEdgeType
	msoSoftEdgeType4              =4          # from enum MsoSoftEdgeType
	msoSoftEdgeType5              =5          # from enum MsoSoftEdgeType
	msoSoftEdgeType6              =6          # from enum MsoSoftEdgeType
	msoSoftEdgeTypeMixed          =-2         # from enum MsoSoftEdgeType
	msoSoftEdgeTypeNone           =0          # from enum MsoSoftEdgeType
	msoSortByFileName             =1          # from enum MsoSortBy
	msoSortByFileType             =3          # from enum MsoSortBy
	msoSortByLastModified         =4          # from enum MsoSortBy
	msoSortByNone                 =5          # from enum MsoSortBy
	msoSortBySize                 =2          # from enum MsoSortBy
	msoSortOrderAscending         =1          # from enum MsoSortOrder
	msoSortOrderDescending        =2          # from enum MsoSortOrder
	msoSyncAvailableAnywhere      =2          # from enum MsoSyncAvailableType
	msoSyncAvailableNone          =0          # from enum MsoSyncAvailableType
	msoSyncAvailableOffline       =1          # from enum MsoSyncAvailableType
	msoSyncCompareAndMerge        =0          # from enum MsoSyncCompareType
	msoSyncCompareSideBySide      =1          # from enum MsoSyncCompareType
	msoSyncConflictClientWins     =0          # from enum MsoSyncConflictResolutionType
	msoSyncConflictMerge          =2          # from enum MsoSyncConflictResolutionType
	msoSyncConflictServerWins     =1          # from enum MsoSyncConflictResolutionType
	msoSyncErrorCouldNotCompare   =13         # from enum MsoSyncErrorType
	msoSyncErrorCouldNotConnect   =2          # from enum MsoSyncErrorType
	msoSyncErrorCouldNotOpen      =11         # from enum MsoSyncErrorType
	msoSyncErrorCouldNotResolve   =14         # from enum MsoSyncErrorType
	msoSyncErrorCouldNotUpdate    =12         # from enum MsoSyncErrorType
	msoSyncErrorFileInUse         =6          # from enum MsoSyncErrorType
	msoSyncErrorFileNotFound      =4          # from enum MsoSyncErrorType
	msoSyncErrorFileTooLarge      =5          # from enum MsoSyncErrorType
	msoSyncErrorNoNetwork         =15         # from enum MsoSyncErrorType
	msoSyncErrorNone              =0          # from enum MsoSyncErrorType
	msoSyncErrorOutOfSpace        =3          # from enum MsoSyncErrorType
	msoSyncErrorUnauthorizedUser  =1          # from enum MsoSyncErrorType
	msoSyncErrorUnknown           =16         # from enum MsoSyncErrorType
	msoSyncErrorUnknownDownload   =10         # from enum MsoSyncErrorType
	msoSyncErrorUnknownUpload     =9          # from enum MsoSyncErrorType
	msoSyncErrorVirusDownload     =8          # from enum MsoSyncErrorType
	msoSyncErrorVirusUpload       =7          # from enum MsoSyncErrorType
	msoSyncEventDownloadFailed    =2          # from enum MsoSyncEventType
	msoSyncEventDownloadInitiated =0          # from enum MsoSyncEventType
	msoSyncEventDownloadNoChange  =6          # from enum MsoSyncEventType
	msoSyncEventDownloadSucceeded =1          # from enum MsoSyncEventType
	msoSyncEventOffline           =7          # from enum MsoSyncEventType
	msoSyncEventUploadFailed      =5          # from enum MsoSyncEventType
	msoSyncEventUploadInitiated   =3          # from enum MsoSyncEventType
	msoSyncEventUploadSucceeded   =4          # from enum MsoSyncEventType
	msoSyncStatusConflict         =4          # from enum MsoSyncStatusType
	msoSyncStatusError            =6          # from enum MsoSyncStatusType
	msoSyncStatusLatest           =1          # from enum MsoSyncStatusType
	msoSyncStatusLocalChanges     =3          # from enum MsoSyncStatusType
	msoSyncStatusNewerAvailable   =2          # from enum MsoSyncStatusType
	msoSyncStatusNoSharedWorkspace=0          # from enum MsoSyncStatusType
	msoSyncStatusNotRoaming       =0          # from enum MsoSyncStatusType
	msoSyncStatusSuspended        =5          # from enum MsoSyncStatusType
	msoSyncVersionLastViewed      =0          # from enum MsoSyncVersionType
	msoSyncVersionServer          =1          # from enum MsoSyncVersionType
	msoTabStopCenter              =2          # from enum MsoTabStopType
	msoTabStopDecimal             =4          # from enum MsoTabStopType
	msoTabStopLeft                =1          # from enum MsoTabStopType
	msoTabStopMixed               =-2         # from enum MsoTabStopType
	msoTabStopRight               =3          # from enum MsoTabStopType
	msoTargetBrowserIE4           =2          # from enum MsoTargetBrowser
	msoTargetBrowserIE5           =3          # from enum MsoTargetBrowser
	msoTargetBrowserIE6           =4          # from enum MsoTargetBrowser
	msoTargetBrowserV3            =0          # from enum MsoTargetBrowser
	msoTargetBrowserV4            =1          # from enum MsoTargetBrowser
	msoAllCaps                    =2          # from enum MsoTextCaps
	msoCapsMixed                  =-2         # from enum MsoTextCaps
	msoNoCaps                     =0          # from enum MsoTextCaps
	msoSmallCaps                  =1          # from enum MsoTextCaps
	msoCaseLower                  =2          # from enum MsoTextChangeCase
	msoCaseSentence               =1          # from enum MsoTextChangeCase
	msoCaseTitle                  =4          # from enum MsoTextChangeCase
	msoCaseToggle                 =5          # from enum MsoTextChangeCase
	msoCaseUpper                  =3          # from enum MsoTextChangeCase
	msoCharWrapMixed              =-2         # from enum MsoTextCharWrap
	msoCustomCharWrap             =3          # from enum MsoTextCharWrap
	msoNoCharWrap                 =0          # from enum MsoTextCharWrap
	msoStandardCharWrap           =1          # from enum MsoTextCharWrap
	msoStrictCharWrap             =2          # from enum MsoTextCharWrap
	msoTextDirectionLeftToRight   =1          # from enum MsoTextDirection
	msoTextDirectionMixed         =-2         # from enum MsoTextDirection
	msoTextDirectionRightToLeft   =2          # from enum MsoTextDirection
	msoTextEffectAlignmentCentered=2          # from enum MsoTextEffectAlignment
	msoTextEffectAlignmentLeft    =1          # from enum MsoTextEffectAlignment
	msoTextEffectAlignmentLetterJustify=4          # from enum MsoTextEffectAlignment
	msoTextEffectAlignmentMixed   =-2         # from enum MsoTextEffectAlignment
	msoTextEffectAlignmentRight   =3          # from enum MsoTextEffectAlignment
	msoTextEffectAlignmentStretchJustify=6          # from enum MsoTextEffectAlignment
	msoTextEffectAlignmentWordJustify=5          # from enum MsoTextEffectAlignment
	msoFontAlignAuto              =0          # from enum MsoTextFontAlign
	msoFontAlignBaseline          =3          # from enum MsoTextFontAlign
	msoFontAlignBottom            =4          # from enum MsoTextFontAlign
	msoFontAlignCenter            =2          # from enum MsoTextFontAlign
	msoFontAlignMixed             =-2         # from enum MsoTextFontAlign
	msoFontAlignTop               =1          # from enum MsoTextFontAlign
	msoTextOrientationDownward    =3          # from enum MsoTextOrientation
	msoTextOrientationHorizontal  =1          # from enum MsoTextOrientation
	msoTextOrientationHorizontalRotatedFarEast=6          # from enum MsoTextOrientation
	msoTextOrientationMixed       =-2         # from enum MsoTextOrientation
	msoTextOrientationUpward      =2          # from enum MsoTextOrientation
	msoTextOrientationVertical    =5          # from enum MsoTextOrientation
	msoTextOrientationVerticalFarEast=4          # from enum MsoTextOrientation
	msoDoubleStrike               =2          # from enum MsoTextStrike
	msoNoStrike                   =0          # from enum MsoTextStrike
	msoSingleStrike               =1          # from enum MsoTextStrike
	msoStrikeMixed                =-2         # from enum MsoTextStrike
	msoTabAlignCenter             =1          # from enum MsoTextTabAlign
	msoTabAlignDecimal            =3          # from enum MsoTextTabAlign
	msoTabAlignLeft               =0          # from enum MsoTextTabAlign
	msoTabAlignMixed              =-2         # from enum MsoTextTabAlign
	msoTabAlignRight              =2          # from enum MsoTextTabAlign
	msoNoUnderline                =0          # from enum MsoTextUnderlineType
	msoUnderlineDashHeavyLine     =8          # from enum MsoTextUnderlineType
	msoUnderlineDashLine          =7          # from enum MsoTextUnderlineType
	msoUnderlineDashLongHeavyLine =10         # from enum MsoTextUnderlineType
	msoUnderlineDashLongLine      =9          # from enum MsoTextUnderlineType
	msoUnderlineDotDashHeavyLine  =12         # from enum MsoTextUnderlineType
	msoUnderlineDotDashLine       =11         # from enum MsoTextUnderlineType
	msoUnderlineDotDotDashHeavyLine=14         # from enum MsoTextUnderlineType
	msoUnderlineDotDotDashLine    =13         # from enum MsoTextUnderlineType
	msoUnderlineDottedHeavyLine   =6          # from enum MsoTextUnderlineType
	msoUnderlineDottedLine        =5          # from enum MsoTextUnderlineType
	msoUnderlineDoubleLine        =3          # from enum MsoTextUnderlineType
	msoUnderlineHeavyLine         =4          # from enum MsoTextUnderlineType
	msoUnderlineMixed             =-2         # from enum MsoTextUnderlineType
	msoUnderlineSingleLine        =2          # from enum MsoTextUnderlineType
	msoUnderlineWavyDoubleLine    =17         # from enum MsoTextUnderlineType
	msoUnderlineWavyHeavyLine     =16         # from enum MsoTextUnderlineType
	msoUnderlineWavyLine          =15         # from enum MsoTextUnderlineType
	msoUnderlineWords             =1          # from enum MsoTextUnderlineType
	msoTextureAlignmentMixed      =-2         # from enum MsoTextureAlignment
	msoTextureBottom              =7          # from enum MsoTextureAlignment
	msoTextureBottomLeft          =6          # from enum MsoTextureAlignment
	msoTextureBottomRight         =8          # from enum MsoTextureAlignment
	msoTextureCenter              =4          # from enum MsoTextureAlignment
	msoTextureLeft                =3          # from enum MsoTextureAlignment
	msoTextureRight               =5          # from enum MsoTextureAlignment
	msoTextureTop                 =1          # from enum MsoTextureAlignment
	msoTextureTopLeft             =0          # from enum MsoTextureAlignment
	msoTextureTopRight            =2          # from enum MsoTextureAlignment
	msoTexturePreset              =1          # from enum MsoTextureType
	msoTextureTypeMixed           =-2         # from enum MsoTextureType
	msoTextureUserDefined         =2          # from enum MsoTextureType
	msoNotThemeColor              =0          # from enum MsoThemeColorIndex
	msoThemeColorAccent1          =5          # from enum MsoThemeColorIndex
	msoThemeColorAccent2          =6          # from enum MsoThemeColorIndex
	msoThemeColorAccent3          =7          # from enum MsoThemeColorIndex
	msoThemeColorAccent4          =8          # from enum MsoThemeColorIndex
	msoThemeColorAccent5          =9          # from enum MsoThemeColorIndex
	msoThemeColorAccent6          =10         # from enum MsoThemeColorIndex
	msoThemeColorBackground1      =14         # from enum MsoThemeColorIndex
	msoThemeColorBackground2      =16         # from enum MsoThemeColorIndex
	msoThemeColorDark1            =1          # from enum MsoThemeColorIndex
	msoThemeColorDark2            =3          # from enum MsoThemeColorIndex
	msoThemeColorFollowedHyperlink=12         # from enum MsoThemeColorIndex
	msoThemeColorHyperlink        =11         # from enum MsoThemeColorIndex
	msoThemeColorLight1           =2          # from enum MsoThemeColorIndex
	msoThemeColorLight2           =4          # from enum MsoThemeColorIndex
	msoThemeColorMixed            =-2         # from enum MsoThemeColorIndex
	msoThemeColorText1            =13         # from enum MsoThemeColorIndex
	msoThemeColorText2            =15         # from enum MsoThemeColorIndex
	msoThemeAccent1               =5          # from enum MsoThemeColorSchemeIndex
	msoThemeAccent2               =6          # from enum MsoThemeColorSchemeIndex
	msoThemeAccent3               =7          # from enum MsoThemeColorSchemeIndex
	msoThemeAccent4               =8          # from enum MsoThemeColorSchemeIndex
	msoThemeAccent5               =9          # from enum MsoThemeColorSchemeIndex
	msoThemeAccent6               =10         # from enum MsoThemeColorSchemeIndex
	msoThemeDark1                 =1          # from enum MsoThemeColorSchemeIndex
	msoThemeDark2                 =3          # from enum MsoThemeColorSchemeIndex
	msoThemeFollowedHyperlink     =12         # from enum MsoThemeColorSchemeIndex
	msoThemeHyperlink             =11         # from enum MsoThemeColorSchemeIndex
	msoThemeLight1                =2          # from enum MsoThemeColorSchemeIndex
	msoThemeLight2                =4          # from enum MsoThemeColorSchemeIndex
	msoCTrue                      =1          # from enum MsoTriState
	msoFalse                      =0          # from enum MsoTriState
	msoTriStateMixed              =-2         # from enum MsoTriState
	msoTriStateToggle             =-3         # from enum MsoTriState
	msoTrue                       =-1         # from enum MsoTriState
	msoAnchorBottom               =4          # from enum MsoVerticalAnchor
	msoAnchorBottomBaseLine       =5          # from enum MsoVerticalAnchor
	msoAnchorMiddle               =3          # from enum MsoVerticalAnchor
	msoAnchorTop                  =1          # from enum MsoVerticalAnchor
	msoAnchorTopBaseline          =2          # from enum MsoVerticalAnchor
	msoVerticalAnchorMixed        =-2         # from enum MsoVerticalAnchor
	msoWarpFormat1                =0          # from enum MsoWarpFormat
	msoWarpFormat10               =9          # from enum MsoWarpFormat
	msoWarpFormat11               =10         # from enum MsoWarpFormat
	msoWarpFormat12               =11         # from enum MsoWarpFormat
	msoWarpFormat13               =12         # from enum MsoWarpFormat
	msoWarpFormat14               =13         # from enum MsoWarpFormat
	msoWarpFormat15               =14         # from enum MsoWarpFormat
	msoWarpFormat16               =15         # from enum MsoWarpFormat
	msoWarpFormat17               =16         # from enum MsoWarpFormat
	msoWarpFormat18               =17         # from enum MsoWarpFormat
	msoWarpFormat19               =18         # from enum MsoWarpFormat
	msoWarpFormat2                =1          # from enum MsoWarpFormat
	msoWarpFormat20               =19         # from enum MsoWarpFormat
	msoWarpFormat21               =20         # from enum MsoWarpFormat
	msoWarpFormat22               =21         # from enum MsoWarpFormat
	msoWarpFormat23               =22         # from enum MsoWarpFormat
	msoWarpFormat24               =23         # from enum MsoWarpFormat
	msoWarpFormat25               =24         # from enum MsoWarpFormat
	msoWarpFormat26               =25         # from enum MsoWarpFormat
	msoWarpFormat27               =26         # from enum MsoWarpFormat
	msoWarpFormat28               =27         # from enum MsoWarpFormat
	msoWarpFormat29               =28         # from enum MsoWarpFormat
	msoWarpFormat3                =2          # from enum MsoWarpFormat
	msoWarpFormat30               =29         # from enum MsoWarpFormat
	msoWarpFormat31               =30         # from enum MsoWarpFormat
	msoWarpFormat32               =31         # from enum MsoWarpFormat
	msoWarpFormat33               =32         # from enum MsoWarpFormat
	msoWarpFormat34               =33         # from enum MsoWarpFormat
	msoWarpFormat35               =34         # from enum MsoWarpFormat
	msoWarpFormat36               =35         # from enum MsoWarpFormat
	msoWarpFormat4                =3          # from enum MsoWarpFormat
	msoWarpFormat5                =4          # from enum MsoWarpFormat
	msoWarpFormat6                =5          # from enum MsoWarpFormat
	msoWarpFormat7                =6          # from enum MsoWarpFormat
	msoWarpFormat8                =7          # from enum MsoWarpFormat
	msoWarpFormat9                =8          # from enum MsoWarpFormat
	msoWarpFormatMixed            =-2         # from enum MsoWarpFormat
	msoWizardActActive            =1          # from enum MsoWizardActType
	msoWizardActInactive          =0          # from enum MsoWizardActType
	msoWizardActResume            =3          # from enum MsoWizardActType
	msoWizardActSuspend           =2          # from enum MsoWizardActType
	msoWizardMsgLocalStateOff     =2          # from enum MsoWizardMsgType
	msoWizardMsgLocalStateOn      =1          # from enum MsoWizardMsgType
	msoWizardMsgResuming          =5          # from enum MsoWizardMsgType
	msoWizardMsgShowHelp          =3          # from enum MsoWizardMsgType
	msoWizardMsgSuspending        =4          # from enum MsoWizardMsgType
	msoBringForward               =2          # from enum MsoZOrderCmd
	msoBringInFrontOfText         =4          # from enum MsoZOrderCmd
	msoBringToFront               =0          # from enum MsoZOrderCmd
	msoSendBackward               =3          # from enum MsoZOrderCmd
	msoSendBehindText             =5          # from enum MsoZOrderCmd
	msoSendToBack                 =1          # from enum MsoZOrderCmd
	RibbonControlSizeLarge        =1          # from enum RibbonControlSize
	RibbonControlSizeRegular      =0          # from enum RibbonControlSize
	sigdetApplicationName         =1          # from enum SignatureDetail
	sigdetApplicationVersion      =2          # from enum SignatureDetail
	sigdetColorDepth              =8          # from enum SignatureDetail
	sigdetDelSuggSigner           =16         # from enum SignatureDetail
	sigdetDelSuggSignerEmail      =20         # from enum SignatureDetail
	sigdetDelSuggSignerEmailSet   =21         # from enum SignatureDetail
	sigdetDelSuggSignerLine2      =18         # from enum SignatureDetail
	sigdetDelSuggSignerLine2Set   =19         # from enum SignatureDetail
	sigdetDelSuggSignerSet        =17         # from enum SignatureDetail
	sigdetDocPreviewImg           =10         # from enum SignatureDetail
	sigdetHashAlgorithm           =14         # from enum SignatureDetail
	sigdetHorizResolution         =6          # from enum SignatureDetail
	sigdetIPCurrentView           =12         # from enum SignatureDetail
	sigdetIPFormHash              =11         # from enum SignatureDetail
	sigdetLocalSigningTime        =0          # from enum SignatureDetail
	sigdetNumberOfMonitors        =5          # from enum SignatureDetail
	sigdetOfficeVersion           =3          # from enum SignatureDetail
	sigdetShouldShowViewWarning   =15         # from enum SignatureDetail
	sigdetSignatureType           =13         # from enum SignatureDetail
	sigdetSignedData              =9          # from enum SignatureDetail
	sigdetVertResolution          =7          # from enum SignatureDetail
	sigdetWindowsVersion          =4          # from enum SignatureDetail
	siglnimgSigned                =4          # from enum SignatureLineImage
	siglnimgSignedInvalid         =3          # from enum SignatureLineImage
	siglnimgSignedValid           =2          # from enum SignatureLineImage
	siglnimgSoftwareRequired      =0          # from enum SignatureLineImage
	siglnimgUnsigned              =1          # from enum SignatureLineImage
	sigprovdetHashAlgorithm       =1          # from enum SignatureProviderDetail
	sigprovdetUIOnly              =2          # from enum SignatureProviderDetail
	sigprovdetUrl                 =0          # from enum SignatureProviderDetail
	sigprovdetUseOfficeStampUI    =4          # from enum SignatureProviderDetail
	sigprovdetUseOfficeUI         =3          # from enum SignatureProviderDetail
	sigtypeMax                    =3          # from enum SignatureType
	sigtypeNonVisible             =1          # from enum SignatureType
	sigtypeSignatureLine          =2          # from enum SignatureType
	sigtypeUnknown                =0          # from enum SignatureType
	xlAxisCrossesAutomatic        =-4105      # from enum XlAxisCrosses
	xlAxisCrossesCustom           =-4114      # from enum XlAxisCrosses
	xlAxisCrossesMaximum          =2          # from enum XlAxisCrosses
	xlAxisCrossesMinimum          =4          # from enum XlAxisCrosses
	xlPrimary                     =1          # from enum XlAxisGroup
	xlSecondary                   =2          # from enum XlAxisGroup
	xlCategory                    =1          # from enum XlAxisType
	xlSeriesAxis                  =3          # from enum XlAxisType
	xlValue                       =2          # from enum XlAxisType
	xlBox                         =0          # from enum XlBarShape
	xlConeToMax                   =5          # from enum XlBarShape
	xlConeToPoint                 =4          # from enum XlBarShape
	xlCylinder                    =3          # from enum XlBarShape
	xlPyramidToMax                =2          # from enum XlBarShape
	xlPyramidToPoint              =1          # from enum XlBarShape
	xlHairline                    =1          # from enum XlBorderWeight
	xlMedium                      =-4138      # from enum XlBorderWeight
	xlThick                       =4          # from enum XlBorderWeight
	xlThin                        =2          # from enum XlBorderWeight
	xlAutomaticScale              =-4105      # from enum XlCategoryType
	xlCategoryScale               =2          # from enum XlCategoryType
	xlTimeScale                   =3          # from enum XlCategoryType
	xlChartElementPositionAutomatic=-4105      # from enum XlChartElementPosition
	xlChartElementPositionCustom  =-4114      # from enum XlChartElementPosition
	xlAxis                        =21         # from enum XlChartItem
	xlAxisTitle                   =17         # from enum XlChartItem
	xlChartArea                   =2          # from enum XlChartItem
	xlChartTitle                  =4          # from enum XlChartItem
	xlCorners                     =6          # from enum XlChartItem
	xlDataLabel                   =0          # from enum XlChartItem
	xlDataTable                   =7          # from enum XlChartItem
	xlDisplayUnitLabel            =30         # from enum XlChartItem
	xlDownBars                    =20         # from enum XlChartItem
	xlDropLines                   =26         # from enum XlChartItem
	xlErrorBars                   =9          # from enum XlChartItem
	xlFloor                       =23         # from enum XlChartItem
	xlHiLoLines                   =25         # from enum XlChartItem
	xlLeaderLines                 =29         # from enum XlChartItem
	xlLegend                      =24         # from enum XlChartItem
	xlLegendEntry                 =12         # from enum XlChartItem
	xlLegendKey                   =13         # from enum XlChartItem
	xlMajorGridlines              =15         # from enum XlChartItem
	xlMinorGridlines              =16         # from enum XlChartItem
	xlNothing                     =28         # from enum XlChartItem
	xlPivotChartDropZone          =32         # from enum XlChartItem
	xlPivotChartFieldButton       =31         # from enum XlChartItem
	xlPlotArea                    =19         # from enum XlChartItem
	xlRadarAxisLabels             =27         # from enum XlChartItem
	xlSeries                      =3          # from enum XlChartItem
	xlSeriesLines                 =22         # from enum XlChartItem
	xlShape                       =14         # from enum XlChartItem
	xlTrendline                   =8          # from enum XlChartItem
	xlUpBars                      =18         # from enum XlChartItem
	xlWalls                       =5          # from enum XlChartItem
	xlXErrorBars                  =10         # from enum XlChartItem
	xlYErrorBars                  =11         # from enum XlChartItem
	xlDownward                    =-4170      # from enum XlChartOrientation
	xlHorizontal                  =-4128      # from enum XlChartOrientation
	xlUpward                      =-4171      # from enum XlChartOrientation
	xlVertical                    =-4166      # from enum XlChartOrientation
	xlStack                       =2          # from enum XlChartPictureType
	xlStackScale                  =3          # from enum XlChartPictureType
	xlStretch                     =1          # from enum XlChartPictureType
	xlSplitByCustomSplit          =4          # from enum XlChartSplitType
	xlSplitByPercentValue         =3          # from enum XlChartSplitType
	xlSplitByPosition             =1          # from enum XlChartSplitType
	xlSplitByValue                =2          # from enum XlChartSplitType
	xl3DArea                      =-4098      # from enum XlChartType
	xl3DAreaStacked               =78         # from enum XlChartType
	xl3DAreaStacked100            =79         # from enum XlChartType
	xl3DBarClustered              =60         # from enum XlChartType
	xl3DBarStacked                =61         # from enum XlChartType
	xl3DBarStacked100             =62         # from enum XlChartType
	xl3DColumn                    =-4100      # from enum XlChartType
	xl3DColumnClustered           =54         # from enum XlChartType
	xl3DColumnStacked             =55         # from enum XlChartType
	xl3DColumnStacked100          =56         # from enum XlChartType
	xl3DLine                      =-4101      # from enum XlChartType
	xl3DPie                       =-4102      # from enum XlChartType
	xl3DPieExploded               =70         # from enum XlChartType
	xlArea                        =1          # from enum XlChartType
	xlAreaStacked                 =76         # from enum XlChartType
	xlAreaStacked100              =77         # from enum XlChartType
	xlBarClustered                =57         # from enum XlChartType
	xlBarOfPie                    =71         # from enum XlChartType
	xlBarStacked                  =58         # from enum XlChartType
	xlBarStacked100               =59         # from enum XlChartType
	xlBubble                      =15         # from enum XlChartType
	xlBubble3DEffect              =87         # from enum XlChartType
	xlColumnClustered             =51         # from enum XlChartType
	xlColumnStacked               =52         # from enum XlChartType
	xlColumnStacked100            =53         # from enum XlChartType
	xlConeBarClustered            =102        # from enum XlChartType
	xlConeBarStacked              =103        # from enum XlChartType
	xlConeBarStacked100           =104        # from enum XlChartType
	xlConeCol                     =105        # from enum XlChartType
	xlConeColClustered            =99         # from enum XlChartType
	xlConeColStacked              =100        # from enum XlChartType
	xlConeColStacked100           =101        # from enum XlChartType
	xlCylinderBarClustered        =95         # from enum XlChartType
	xlCylinderBarStacked          =96         # from enum XlChartType
	xlCylinderBarStacked100       =97         # from enum XlChartType
	xlCylinderCol                 =98         # from enum XlChartType
	xlCylinderColClustered        =92         # from enum XlChartType
	xlCylinderColStacked          =93         # from enum XlChartType
	xlCylinderColStacked100       =94         # from enum XlChartType
	xlDoughnut                    =-4120      # from enum XlChartType
	xlDoughnutExploded            =80         # from enum XlChartType
	xlLine                        =4          # from enum XlChartType
	xlLineMarkers                 =65         # from enum XlChartType
	xlLineMarkersStacked          =66         # from enum XlChartType
	xlLineMarkersStacked100       =67         # from enum XlChartType
	xlLineStacked                 =63         # from enum XlChartType
	xlLineStacked100              =64         # from enum XlChartType
	xlPie                         =5          # from enum XlChartType
	xlPieExploded                 =69         # from enum XlChartType
	xlPieOfPie                    =68         # from enum XlChartType
	xlPyramidBarClustered         =109        # from enum XlChartType
	xlPyramidBarStacked           =110        # from enum XlChartType
	xlPyramidBarStacked100        =111        # from enum XlChartType
	xlPyramidCol                  =112        # from enum XlChartType
	xlPyramidColClustered         =106        # from enum XlChartType
	xlPyramidColStacked           =107        # from enum XlChartType
	xlPyramidColStacked100        =108        # from enum XlChartType
	xlRadar                       =-4151      # from enum XlChartType
	xlRadarFilled                 =82         # from enum XlChartType
	xlRadarMarkers                =81         # from enum XlChartType
	xlStockHLC                    =88         # from enum XlChartType
	xlStockOHLC                   =89         # from enum XlChartType
	xlStockVHLC                   =90         # from enum XlChartType
	xlStockVOHLC                  =91         # from enum XlChartType
	xlSurface                     =83         # from enum XlChartType
	xlSurfaceTopView              =85         # from enum XlChartType
	xlSurfaceTopViewWireframe     =86         # from enum XlChartType
	xlSurfaceWireframe            =84         # from enum XlChartType
	xlXYScatter                   =-4169      # from enum XlChartType
	xlXYScatterLines              =74         # from enum XlChartType
	xlXYScatterLinesNoMarkers     =75         # from enum XlChartType
	xlXYScatterSmooth             =72         # from enum XlChartType
	xlXYScatterSmoothNoMarkers    =73         # from enum XlChartType
	xlColorIndexAutomatic         =-4105      # from enum XlColorIndex
	xlColorIndexNone              =-4142      # from enum XlColorIndex
	xl3DBar                       =-4099      # from enum XlConstants
	xl3DSurface                   =-4103      # from enum XlConstants
	xlAutomatic                   =-4105      # from enum XlConstants
	xlBar                         =2          # from enum XlConstants
	xlColumn                      =3          # from enum XlConstants
	xlCombination                 =-4111      # from enum XlConstants
	xlCustom                      =-4114      # from enum XlConstants
	xlDefaultAutoFormat           =-1         # from enum XlConstants
	xlNone                        =-4142      # from enum XlConstants
	xlLabelPositionAbove          =0          # from enum XlDataLabelPosition
	xlLabelPositionBelow          =1          # from enum XlDataLabelPosition
	xlLabelPositionBestFit        =5          # from enum XlDataLabelPosition
	xlLabelPositionCenter         =-4108      # from enum XlDataLabelPosition
	xlLabelPositionCustom         =7          # from enum XlDataLabelPosition
	xlLabelPositionInsideBase     =4          # from enum XlDataLabelPosition
	xlLabelPositionInsideEnd      =3          # from enum XlDataLabelPosition
	xlLabelPositionLeft           =-4131      # from enum XlDataLabelPosition
	xlLabelPositionMixed          =6          # from enum XlDataLabelPosition
	xlLabelPositionOutsideEnd     =2          # from enum XlDataLabelPosition
	xlLabelPositionRight          =-4152      # from enum XlDataLabelPosition
	xlDataLabelsShowBubbleSizes   =6          # from enum XlDataLabelsType
	xlDataLabelsShowLabel         =4          # from enum XlDataLabelsType
	xlDataLabelsShowLabelAndPercent=5          # from enum XlDataLabelsType
	xlDataLabelsShowNone          =-4142      # from enum XlDataLabelsType
	xlDataLabelsShowPercent       =3          # from enum XlDataLabelsType
	xlDataLabelsShowValue         =2          # from enum XlDataLabelsType
	xlInterpolated                =3          # from enum XlDisplayBlanksAs
	xlNotPlotted                  =1          # from enum XlDisplayBlanksAs
	xlZero                        =2          # from enum XlDisplayBlanksAs
	xlDisplayUnitCustom           =-4114      # from enum XlDisplayUnit
	xlDisplayUnitNone             =-4142      # from enum XlDisplayUnit
	xlHundredMillions             =-8         # from enum XlDisplayUnit
	xlHundredThousands            =-5         # from enum XlDisplayUnit
	xlHundreds                    =-2         # from enum XlDisplayUnit
	xlMillionMillions             =-10        # from enum XlDisplayUnit
	xlMillions                    =-6         # from enum XlDisplayUnit
	xlTenMillions                 =-7         # from enum XlDisplayUnit
	xlTenThousands                =-4         # from enum XlDisplayUnit
	xlThousandMillions            =-9         # from enum XlDisplayUnit
	xlThousands                   =-3         # from enum XlDisplayUnit
	xlCap                         =1          # from enum XlEndStyleCap
	xlNoCap                       =2          # from enum XlEndStyleCap
	xlChartX                      =-4168      # from enum XlErrorBarDirection
	xlChartY                      =1          # from enum XlErrorBarDirection
	xlErrorBarIncludeBoth         =1          # from enum XlErrorBarInclude
	xlErrorBarIncludeMinusValues  =3          # from enum XlErrorBarInclude
	xlErrorBarIncludeNone         =-4142      # from enum XlErrorBarInclude
	xlErrorBarIncludePlusValues   =2          # from enum XlErrorBarInclude
	xlErrorBarTypeCustom          =-4114      # from enum XlErrorBarType
	xlErrorBarTypeFixedValue      =1          # from enum XlErrorBarType
	xlErrorBarTypePercent         =2          # from enum XlErrorBarType
	xlErrorBarTypeStDev           =-4155      # from enum XlErrorBarType
	xlErrorBarTypeStError         =4          # from enum XlErrorBarType
	xlHAlignCenter                =-4108      # from enum XlHAlign
	xlHAlignCenterAcrossSelection =7          # from enum XlHAlign
	xlHAlignDistributed           =-4117      # from enum XlHAlign
	xlHAlignFill                  =5          # from enum XlHAlign
	xlHAlignGeneral               =1          # from enum XlHAlign
	xlHAlignJustify               =-4130      # from enum XlHAlign
	xlHAlignLeft                  =-4131      # from enum XlHAlign
	xlHAlignRight                 =-4152      # from enum XlHAlign
	xlLegendPositionBottom        =-4107      # from enum XlLegendPosition
	xlLegendPositionCorner        =2          # from enum XlLegendPosition
	xlLegendPositionCustom        =-4161      # from enum XlLegendPosition
	xlLegendPositionLeft          =-4131      # from enum XlLegendPosition
	xlLegendPositionRight         =-4152      # from enum XlLegendPosition
	xlLegendPositionTop           =-4160      # from enum XlLegendPosition
	xlMarkerStyleAutomatic        =-4105      # from enum XlMarkerStyle
	xlMarkerStyleCircle           =8          # from enum XlMarkerStyle
	xlMarkerStyleDash             =-4115      # from enum XlMarkerStyle
	xlMarkerStyleDiamond          =2          # from enum XlMarkerStyle
	xlMarkerStyleDot              =-4118      # from enum XlMarkerStyle
	xlMarkerStyleNone             =-4142      # from enum XlMarkerStyle
	xlMarkerStylePicture          =-4147      # from enum XlMarkerStyle
	xlMarkerStylePlus             =9          # from enum XlMarkerStyle
	xlMarkerStyleSquare           =1          # from enum XlMarkerStyle
	xlMarkerStyleStar             =5          # from enum XlMarkerStyle
	xlMarkerStyleTriangle         =3          # from enum XlMarkerStyle
	xlMarkerStyleX                =-4168      # from enum XlMarkerStyle
	xlColumnField                 =2          # from enum XlPivotFieldOrientation
	xlDataField                   =4          # from enum XlPivotFieldOrientation
	xlHidden                      =0          # from enum XlPivotFieldOrientation
	xlPageField                   =3          # from enum XlPivotFieldOrientation
	xlRowField                    =1          # from enum XlPivotFieldOrientation
	xlContext                     =-5002      # from enum XlReadingOrder
	xlLTR                         =-5003      # from enum XlReadingOrder
	xlRTL                         =-5004      # from enum XlReadingOrder
	xlColumns                     =2          # from enum XlRowCol
	xlRows                        =1          # from enum XlRowCol
	xlScaleLinear                 =-4132      # from enum XlScaleType
	xlScaleLogarithmic            =-4133      # from enum XlScaleType
	xlSizeIsArea                  =1          # from enum XlSizeRepresents
	xlSizeIsWidth                 =2          # from enum XlSizeRepresents
	xlTickLabelOrientationAutomatic=-4105      # from enum XlTickLabelOrientation
	xlTickLabelOrientationDownward=-4170      # from enum XlTickLabelOrientation
	xlTickLabelOrientationHorizontal=-4128      # from enum XlTickLabelOrientation
	xlTickLabelOrientationUpward  =-4171      # from enum XlTickLabelOrientation
	xlTickLabelOrientationVertical=-4166      # from enum XlTickLabelOrientation
	xlTickLabelPositionHigh       =-4127      # from enum XlTickLabelPosition
	xlTickLabelPositionLow        =-4134      # from enum XlTickLabelPosition
	xlTickLabelPositionNextToAxis =4          # from enum XlTickLabelPosition
	xlTickLabelPositionNone       =-4142      # from enum XlTickLabelPosition
	xlTickMarkCross               =4          # from enum XlTickMark
	xlTickMarkInside              =2          # from enum XlTickMark
	xlTickMarkNone                =-4142      # from enum XlTickMark
	xlTickMarkOutside             =3          # from enum XlTickMark
	xlDays                        =0          # from enum XlTimeUnit
	xlMonths                      =1          # from enum XlTimeUnit
	xlYears                       =2          # from enum XlTimeUnit
	xlExponential                 =5          # from enum XlTrendlineType
	xlLinear                      =-4132      # from enum XlTrendlineType
	xlLogarithmic                 =-4133      # from enum XlTrendlineType
	xlMovingAvg                   =6          # from enum XlTrendlineType
	xlPolynomial                  =3          # from enum XlTrendlineType
	xlPower                       =4          # from enum XlTrendlineType
	xlUnderlineStyleDouble        =-4119      # from enum XlUnderlineStyle
	xlUnderlineStyleDoubleAccounting=5          # from enum XlUnderlineStyle
	xlUnderlineStyleNone          =-4142      # from enum XlUnderlineStyle
	xlUnderlineStyleSingle        =2          # from enum XlUnderlineStyle
	xlUnderlineStyleSingleAccounting=4          # from enum XlUnderlineStyle
	xlVAlignBottom                =-4107      # from enum XlVAlign
	xlVAlignCenter                =-4108      # from enum XlVAlign
	xlVAlignDistributed           =-4117      # from enum XlVAlign
	xlVAlignJustify               =-4130      # from enum XlVAlign
	xlVAlignTop                   =-4160      # from enum XlVAlign

RecordMap = {
}

CLSIDToClassMap = {}
CLSIDToPackageMap = {
	'{000C0340-0000-0000-C000-000000000046}' : u'Scripts',
	'{000C0341-0000-0000-C000-000000000046}' : u'Script',
	'{2DF8D04D-5BFA-101B-BDE5-00AA0044DE52}' : u'DocumentProperties',
	'{2DF8D04E-5BFA-101B-BDE5-00AA0044DE52}' : u'DocumentProperty',
	'{000C0351-0000-0000-C000-000000000046}' : u'_CommandBarButtonEvents',
	'{000C0352-0000-0000-C000-000000000046}' : u'_CommandBarsEvents',
	'{000C0353-0000-0000-C000-000000000046}' : u'LanguageSettings',
	'{000C0354-0000-0000-C000-000000000046}' : u'_CommandBarComboBoxEvents',
	'{000C0356-0000-0000-C000-000000000046}' : u'HTMLProject',
	'{000C0357-0000-0000-C000-000000000046}' : u'HTMLProjectItems',
	'{000C0358-0000-0000-C000-000000000046}' : u'HTMLProjectItem',
	'{000C0359-0000-0000-C000-000000000046}' : u'IMsoDispCagNotifySink',
	'{000C035A-0000-0000-C000-000000000046}' : u'MsoDebugOptions',
	'{000C1720-0000-0000-C000-000000000046}' : u'IMsoDataLabel',
	'{000C0360-0000-0000-C000-000000000046}' : u'AnswerWizard',
	'{000C0361-0000-0000-C000-000000000046}' : u'AnswerWizardFiles',
	'{000C0362-0000-0000-C000-000000000046}' : u'FileDialog',
	'{000C0363-0000-0000-C000-000000000046}' : u'FileDialogSelectedItems',
	'{000C0364-0000-0000-C000-000000000046}' : u'FileDialogFilter',
	'{000C0365-0000-0000-C000-000000000046}' : u'FileDialogFilters',
	'{000C0366-0000-0000-C000-000000000046}' : u'SearchScopes',
	'{000C0367-0000-0000-C000-000000000046}' : u'SearchScope',
	'{000C0368-0000-0000-C000-000000000046}' : u'ScopeFolder',
	'{000C0369-0000-0000-C000-000000000046}' : u'ScopeFolders',
	'{000C036A-0000-0000-C000-000000000046}' : u'SearchFolders',
	'{000C036C-0000-0000-C000-000000000046}' : u'FileTypes',
	'{000C036D-0000-0000-C000-000000000046}' : u'IMsoDiagram',
	'{000C036E-0000-0000-C000-000000000046}' : u'DiagramNodes',
	'{000C036F-0000-0000-C000-000000000046}' : u'DiagramNodeChildren',
	'{000C0370-0000-0000-C000-000000000046}' : u'DiagramNode',
	'{000C0371-0000-0000-C000-000000000046}' : u'CanvasShapes',
	'{000C0372-0000-0000-C000-000000000046}' : u'IMsoEServicesDialog',
	'{000C0373-0000-0000-C000-000000000046}' : u'WebComponentProperties',
	'{000C0375-0000-0000-C000-000000000046}' : u'UserPermission',
	'{000C0376-0000-0000-C000-000000000046}' : u'Permission',
	'{000C0377-0000-0000-C000-000000000046}' : u'SmartDocument',
	'{000C0379-0000-0000-C000-000000000046}' : u'SharedWorkspaceTask',
	'{000C037A-0000-0000-C000-000000000046}' : u'SharedWorkspaceTasks',
	'{000C037B-0000-0000-C000-000000000046}' : u'SharedWorkspaceFile',
	'{000C037C-0000-0000-C000-000000000046}' : u'SharedWorkspaceFiles',
	'{000C037D-0000-0000-C000-000000000046}' : u'SharedWorkspaceFolder',
	'{000C037E-0000-0000-C000-000000000046}' : u'SharedWorkspaceFolders',
	'{000C037F-0000-0000-C000-000000000046}' : u'SharedWorkspaceLink',
	'{000C0380-0000-0000-C000-000000000046}' : u'SharedWorkspaceLinks',
	'{000C0381-0000-0000-C000-000000000046}' : u'SharedWorkspaceMember',
	'{000C0382-0000-0000-C000-000000000046}' : u'SharedWorkspaceMembers',
	'{000C0385-0000-0000-C000-000000000046}' : u'SharedWorkspace',
	'{000C0386-0000-0000-C000-000000000046}' : u'Sync',
	'{000C0387-0000-0000-C000-000000000046}' : u'DocumentLibraryVersion',
	'{000C0388-0000-0000-C000-000000000046}' : u'DocumentLibraryVersions',
	'{000C0389-0000-0000-C000-000000000046}' : u'MsoDebugOptions_UTManager',
	'{000C038A-0000-0000-C000-000000000046}' : u'MsoDebugOptions_UTs',
	'{000C038B-0000-0000-C000-000000000046}' : u'MsoDebugOptions_UT',
	'{000C038C-0000-0000-C000-000000000046}' : u'MsoDebugOptions_UTRunResult',
	'{000C038E-0000-0000-C000-000000000046}' : u'MetaProperties',
	'{000C038F-0000-0000-C000-000000000046}' : u'MetaProperty',
	'{000C0390-0000-0000-C000-000000000046}' : u'ServerPolicy',
	'{000C0391-0000-0000-C000-000000000046}' : u'PolicyItem',
	'{000C0392-0000-0000-C000-000000000046}' : u'DocumentInspectors',
	'{000C0393-0000-0000-C000-000000000046}' : u'DocumentInspector',
	'{000C0395-0000-0000-C000-000000000046}' : u'IRibbonControl',
	'{00194002-D9C3-11D3-8D59-0050048384E3}' : u'ILicAgent',
	'{000C0397-0000-0000-C000-000000000046}' : u'TextRange2',
	'{000C0398-0000-0000-C000-000000000046}' : u'TextFrame2',
	'{000C0399-0000-0000-C000-000000000046}' : u'ParagraphFormat2',
	'{000C039A-0000-0000-C000-000000000046}' : u'Font2',
	'{000C03A0-0000-0000-C000-000000000046}' : u'OfficeTheme',
	'{000C03A1-0000-0000-C000-000000000046}' : u'ThemeColor',
	'{000C03A2-0000-0000-C000-000000000046}' : u'ThemeColorScheme',
	'{000C03A3-0000-0000-C000-000000000046}' : u'ThemeFont',
	'{000C03A4-0000-0000-C000-000000000046}' : u'ThemeFonts',
	'{000C03A5-0000-0000-C000-000000000046}' : u'ThemeFontScheme',
	'{000C03A6-0000-0000-C000-000000000046}' : u'ThemeEffectScheme',
	'{000C03A7-0000-0000-C000-000000000046}' : u'IRibbonUI',
	'{919AA22C-B9AD-11D3-8D59-0050048384E3}' : u'ILicValidator',
	'{000C170A-0000-0000-C000-000000000046}' : u'SeriesCollection',
	'{000C03B2-0000-0000-C000-000000000046}' : u'TextColumn2',
	'{000C03B9-0000-0000-C000-000000000046}' : u'BulletFormat2',
	'{000C03BA-0000-0000-C000-000000000046}' : u'TabStops2',
	'{000C03BB-0000-0000-C000-000000000046}' : u'TabStop2',
	'{000C03BC-0000-0000-C000-000000000046}' : u'SoftEdgeFormat',
	'{000C03BD-0000-0000-C000-000000000046}' : u'GlowFormat',
	'{000C03BE-0000-0000-C000-000000000046}' : u'ReflectionFormat',
	'{000C03BF-0000-0000-C000-000000000046}' : u'GradientStop',
	'{000C03C0-0000-0000-C000-000000000046}' : u'GradientStops',
	'{000CD100-0000-0000-C000-000000000046}' : u'WebComponent',
	'{4291224C-DEFE-485B-8E69-6CF8AA85CB76}' : u'IAssistance',
	'{000C03C3-0000-0000-C000-000000000046}' : u'RulerLevel2',
	'{000C03C4-0000-0000-C000-000000000046}' : u'IBlogExtensibility',
	'{000C03C5-0000-0000-C000-000000000046}' : u'IBlogPictureExtensibility',
	'{000CD101-0000-0000-C000-000000000046}' : u'WebComponentWindowExternal',
	'{000CD102-0000-0000-C000-000000000046}' : u'WebComponentFormat',
	'{000CDB03-0000-0000-C000-000000000046}' : u'CustomXMLNodes',
	'{ABFA087C-F703-4D53-946E-37FF82B2C994}' : u'IMsoAxisTitle',
	'{000CDB04-0000-0000-C000-000000000046}' : u'CustomXMLNode',
	'{000C170C-0000-0000-C000-000000000046}' : u'ChartPoint',
	'{000CDB08-0000-0000-C000-000000000046}' : u'CustomXMLPart',
	'{000C1709-0000-0000-C000-000000000046}' : u'IMsoChart',
	'{C5771BE5-A188-466B-AB31-00A6A32B1B1C}' : u'CustomTaskPane',
	'{000CDB0A-0000-0000-C000-000000000046}' : u'ICustomXMLPartsEvents',
	'{000C170B-0000-0000-C000-000000000046}' : u'IMsoSeries',
	'{000CDB0C-0000-0000-C000-000000000046}' : u'CustomXMLParts',
	'{000C170D-0000-0000-C000-000000000046}' : u'Points',
	'{000CDB0D-0000-0000-C000-000000000046}' : u'CustomXMLSchemaCollection',
	'{000C0410-0000-0000-C000-000000000046}' : u'SignatureSet',
	'{000C0411-0000-0000-C000-000000000046}' : u'Signature',
	'{000CDB0E-0000-0000-C000-000000000046}' : u'CustomXMLValidationError',
	'{0006F01A-0000-0000-C000-000000000046}' : u'MsoEnvelope',
	'{000C170F-0000-0000-C000-000000000046}' : u'IMsoChartTitle',
	'{000CDB10-0000-0000-C000-000000000046}' : u'CustomXMLPrefixMapping',
	'{A98639A1-CB0C-4A5C-A511-96547F752ACD}' : u'IMsoHyperlinks',
	'{000C1730-0000-0000-C000-000000000046}' : u'IMsoChartFormat',
	'{000C1711-0000-0000-C000-000000000046}' : u'IMsoDataTable',
	'{000C170E-0000-0000-C000-000000000046}' : u'IMsoTrendline',
	'{000C1712-0000-0000-C000-000000000046}' : u'Axes',
	'{6EA00553-9439-4D5A-B1E6-DC15A54DA8B2}' : u'IMsoDisplayUnitLabel',
	'{000C0313-0000-0000-C000-000000000046}' : u'ConnectorFormat',
	'{000CD809-0000-0000-C000-000000000046}' : u'EncryptionProvider',
	'{000C0314-0000-0000-C000-000000000046}' : u'FillFormat',
	'{000C0396-0000-0000-C000-000000000046}' : u'IRibbonExtensibility',
	'{000C1731-0000-0000-C000-000000000046}' : u'IMsoCharacters',
	'{000C1716-0000-0000-C000-000000000046}' : u'IMsoFloor',
	'{000C03C1-0000-0000-C000-000000000046}' : u'Ruler2',
	'{000C1717-0000-0000-C000-000000000046}' : u'IMsoBorder',
	'{000C03C2-0000-0000-C000-000000000046}' : u'RulerLevels2',
	'{000C1718-0000-0000-C000-000000000046}' : u'ChartFont',
	'{000C1719-0000-0000-C000-000000000046}' : u'LegendEntries',
	'{000C171A-0000-0000-C000-000000000046}' : u'LegendEntry',
	'{000C171B-0000-0000-C000-000000000046}' : u'IMsoInterior',
	'{000C1710-0000-0000-C000-000000000046}' : u'IMsoLegend',
	'{000C171C-0000-0000-C000-000000000046}' : u'ChartFillFormat',
	'{000C171D-0000-0000-C000-000000000046}' : u'ChartColorFormat',
	'{000C171E-0000-0000-C000-000000000046}' : u'IMsoLegendKey',
	'{000C171F-0000-0000-C000-000000000046}' : u'IMsoDataLabels',
	'{000C1722-0000-0000-C000-000000000046}' : u'Trendlines',
	'{000C1721-0000-0000-C000-000000000046}' : u'IMsoErrorBars',
	'{55F88890-7708-11D1-ACEB-006008961DA5}' : u'ICommandBarButtonEvents',
	'{55F88891-7708-11D1-ACEB-006008961DA5}' : u'CommandBarButton',
	'{55F88892-7708-11D1-ACEB-006008961DA5}' : u'ICommandBarsEvents',
	'{55F88893-7708-11D1-ACEB-006008961DA5}' : u'CommandBars',
	'{55F88896-7708-11D1-ACEB-006008961DA5}' : u'ICommandBarComboBoxEvents',
	'{55F88897-7708-11D1-ACEB-006008961DA5}' : u'CommandBarComboBox',
	'{000CD900-0000-0000-C000-000000000046}' : u'WorkflowTask',
	'{000C1724-0000-0000-C000-000000000046}' : u'IMsoPlotArea',
	'{000CD6A1-0000-0000-C000-000000000046}' : u'SignatureSetup',
	'{000CD6A2-0000-0000-C000-000000000046}' : u'SignatureInfo',
	'{000CD6A3-0000-0000-C000-000000000046}' : u'SignatureProvider',
	'{000CDB00-0000-0000-C000-000000000046}' : u'CustomXMLPrefixMappings',
	'{000C1726-0000-0000-C000-000000000046}' : u'IMsoTickLabels',
	'{4CAC6328-B9B0-11D3-8D59-0050048384E3}' : u'ILicWizExternal',
	'{000672AC-0000-0000-C000-000000000046}' : u'IMsoEnvelopeVB',
	'{000672AD-0000-0000-C000-000000000046}' : u'IMsoEnvelopeVBEvents',
	'{000C1728-0000-0000-C000-000000000046}' : u'IMsoChartArea',
	'{000CD901-0000-0000-C000-000000000046}' : u'WorkflowTasks',
	'{000C1713-0000-0000-C000-000000000046}' : u'IMsoAxis',
	'{000CDB01-0000-0000-C000-000000000046}' : u'CustomXMLSchema',
	'{000C172C-0001-0000-C000-000000000046}' : u'IMsoDropLines',
	'{000CD902-0000-0000-C000-000000000046}' : u'WorkflowTemplate',
	'{8A64A872-FC6B-4D4A-926E-3A3689562C1C}' : u'CustomTaskPaneEvents',
	'{000C172E-0000-0000-C000-000000000046}' : u'IMsoHiLoLines',
	'{000C1714-0000-0000-C000-000000000046}' : u'IMsoCorners',
	'{618736E0-3C3D-11CF-810C-00AA00389B71}' : u'IAccessible',
	'{000CDB02-0000-0000-C000-000000000046}' : u'_CustomXMLSchemaCollection',
	'{000C1530-0000-0000-C000-000000000046}' : u'OfficeDataSourceObject',
	'{000C1531-0000-0000-C000-000000000046}' : u'ODSOColumn',
	'{000C1532-0000-0000-C000-000000000046}' : u'ODSOColumns',
	'{000CDB06-0000-0000-C000-000000000046}' : u'ICustomXMLPartEvents',
	'{000C1533-0000-0000-C000-000000000046}' : u'ODSOFilter',
	'{000C1534-0000-0000-C000-000000000046}' : u'ODSOFilters',
	'{000C1715-0000-0000-C000-000000000046}' : u'IMsoWalls',
	'{000C0300-0000-0000-C000-000000000046}' : u'_IMsoDispObj',
	'{000C0301-0000-0000-C000-000000000046}' : u'_IMsoOleAccDispObj',
	'{000C0302-0000-0000-C000-000000000046}' : u'_CommandBars',
	'{000CD903-0000-0000-C000-000000000046}' : u'WorkflowTemplates',
	'{000C0304-0000-0000-C000-000000000046}' : u'CommandBar',
	'{000CDB05-0000-0000-C000-000000000046}' : u'_CustomXMLPart',
	'{000C0306-0000-0000-C000-000000000046}' : u'CommandBarControls',
	'{000CDB07-0000-0000-C000-000000000046}' : u'_CustomXMLPartEvents',
	'{000C0308-0000-0000-C000-000000000046}' : u'CommandBarControl',
	'{000CDB09-0000-0000-C000-000000000046}' : u'_CustomXMLParts',
	'{000C030A-0000-0000-C000-000000000046}' : u'CommandBarPopup',
	'{000CDB0B-0000-0000-C000-000000000046}' : u'_CustomXMLPartsEvents',
	'{000C030C-0000-0000-C000-000000000046}' : u'_CommandBarComboBox',
	'{000C030D-0000-0000-C000-000000000046}' : u'_CommandBarActiveX',
	'{000C030E-0000-0000-C000-000000000046}' : u'_CommandBarButton',
	'{000CDB0F-0000-0000-C000-000000000046}' : u'CustomXMLValidationErrors',
	'{000C0310-0000-0000-C000-000000000046}' : u'Adjustments',
	'{000C0311-0000-0000-C000-000000000046}' : u'CalloutFormat',
	'{000C0312-0000-0000-C000-000000000046}' : u'ColorFormat',
	'{000C0913-0000-0000-C000-000000000046}' : u'WebPageFont',
	'{000C0914-0000-0000-C000-000000000046}' : u'WebPageFonts',
	'{000C0315-0000-0000-C000-000000000046}' : u'FreeformBuilder',
	'{000C0316-0000-0000-C000-000000000046}' : u'GroupShapes',
	'{000C0317-0000-0000-C000-000000000046}' : u'LineFormat',
	'{000C0318-0000-0000-C000-000000000046}' : u'ShapeNode',
	'{000C0319-0000-0000-C000-000000000046}' : u'ShapeNodes',
	'{000C031A-0000-0000-C000-000000000046}' : u'PictureFormat',
	'{000C031B-0000-0000-C000-000000000046}' : u'ShadowFormat',
	'{000C031C-0000-0000-C000-000000000046}' : u'Shape',
	'{000C031D-0000-0000-C000-000000000046}' : u'ShapeRange',
	'{000C031E-0000-0000-C000-000000000046}' : u'Shapes',
	'{000C031F-0000-0000-C000-000000000046}' : u'TextEffectFormat',
	'{000C0320-0000-0000-C000-000000000046}' : u'TextFrame',
	'{000C0321-0000-0000-C000-000000000046}' : u'ThreeDFormat',
	'{000C0322-0000-0000-C000-000000000046}' : u'Assistant',
	'{000C1723-0000-0000-C000-000000000046}' : u'IMsoLeaderLines',
	'{000C0324-0000-0000-C000-000000000046}' : u'Balloon',
	'{000C1725-0000-0000-C000-000000000046}' : u'GridLines',
	'{000C0326-0000-0000-C000-000000000046}' : u'BalloonCheckboxes',
	'{000C1727-0000-0000-C000-000000000046}' : u'IMsoChartGroup',
	'{000C0328-0000-0000-C000-000000000046}' : u'BalloonCheckbox',
	'{000C1729-0000-0000-C000-000000000046}' : u'IMsoSeriesLines',
	'{000C172A-0000-0000-C000-000000000046}' : u'IMsoUpBars',
	'{000C172B-0000-0000-C000-000000000046}' : u'ChartGroups',
	'{000C172D-0000-0000-C000-000000000046}' : u'IMsoDownBars',
	'{000C032E-0000-0000-C000-000000000046}' : u'BalloonLabels',
	'{000C172F-0000-0000-C000-000000000046}' : u'IMsoChartData',
	'{000C0330-0000-0000-C000-000000000046}' : u'BalloonLabel',
	'{000C0331-0000-0000-C000-000000000046}' : u'FoundFiles',
	'{000C0332-0000-0000-C000-000000000046}' : u'FileSearch',
	'{000C0333-0000-0000-C000-000000000046}' : u'PropertyTest',
	'{000C0334-0000-0000-C000-000000000046}' : u'PropertyTests',
	'{000C0936-0000-0000-C000-000000000046}' : u'NewFile',
	'{000C0337-0000-0000-C000-000000000046}' : u'IFind',
	'{000C0338-0000-0000-C000-000000000046}' : u'IFoundFiles',
	'{000C0339-0000-0000-C000-000000000046}' : u'COMAddIns',
	'{000C033A-0000-0000-C000-000000000046}' : u'COMAddIn',
	'{000C033B-0000-0000-C000-000000000046}' : u'_CustomTaskPane',
	'{000C033C-0000-0000-C000-000000000046}' : u'_CustomTaskPaneEvents',
	'{000C033D-0000-0000-C000-000000000046}' : u'ICTPFactory',
	'{000C033E-0000-0000-C000-000000000046}' : u'ICustomTaskPaneConsumer',
}
VTablesToClassMap = {}
VTablesToPackageMap = {
	'{000C0340-0000-0000-C000-000000000046}' : 'Scripts',
	'{000C0341-0000-0000-C000-000000000046}' : 'Script',
	'{000C0396-0000-0000-C000-000000000046}' : 'IRibbonExtensibility',
	'{000C0353-0000-0000-C000-000000000046}' : 'LanguageSettings',
	'{000C0356-0000-0000-C000-000000000046}' : 'HTMLProject',
	'{000C0357-0000-0000-C000-000000000046}' : 'HTMLProjectItems',
	'{000C0358-0000-0000-C000-000000000046}' : 'HTMLProjectItem',
	'{000C0359-0000-0000-C000-000000000046}' : 'IMsoDispCagNotifySink',
	'{000C035A-0000-0000-C000-000000000046}' : 'MsoDebugOptions',
	'{000C0360-0000-0000-C000-000000000046}' : 'AnswerWizard',
	'{000C0361-0000-0000-C000-000000000046}' : 'AnswerWizardFiles',
	'{000C0362-0000-0000-C000-000000000046}' : 'FileDialog',
	'{000C0363-0000-0000-C000-000000000046}' : 'FileDialogSelectedItems',
	'{000C0364-0000-0000-C000-000000000046}' : 'FileDialogFilter',
	'{000C0365-0000-0000-C000-000000000046}' : 'FileDialogFilters',
	'{000C0366-0000-0000-C000-000000000046}' : 'SearchScopes',
	'{000C0367-0000-0000-C000-000000000046}' : 'SearchScope',
	'{000C0368-0000-0000-C000-000000000046}' : 'ScopeFolder',
	'{000C0369-0000-0000-C000-000000000046}' : 'ScopeFolders',
	'{000C036A-0000-0000-C000-000000000046}' : 'SearchFolders',
	'{000C036C-0000-0000-C000-000000000046}' : 'FileTypes',
	'{000C036D-0000-0000-C000-000000000046}' : 'IMsoDiagram',
	'{000C036E-0000-0000-C000-000000000046}' : 'DiagramNodes',
	'{000C036F-0000-0000-C000-000000000046}' : 'DiagramNodeChildren',
	'{000C0370-0000-0000-C000-000000000046}' : 'DiagramNode',
	'{000C0371-0000-0000-C000-000000000046}' : 'CanvasShapes',
	'{000C0372-0000-0000-C000-000000000046}' : 'IMsoEServicesDialog',
	'{000C0373-0000-0000-C000-000000000046}' : 'WebComponentProperties',
	'{000C0375-0000-0000-C000-000000000046}' : 'UserPermission',
	'{000C0376-0000-0000-C000-000000000046}' : 'Permission',
	'{000C0377-0000-0000-C000-000000000046}' : 'SmartDocument',
	'{000C0379-0000-0000-C000-000000000046}' : 'SharedWorkspaceTask',
	'{000C037A-0000-0000-C000-000000000046}' : 'SharedWorkspaceTasks',
	'{000C037B-0000-0000-C000-000000000046}' : 'SharedWorkspaceFile',
	'{000C037C-0000-0000-C000-000000000046}' : 'SharedWorkspaceFiles',
	'{000C037D-0000-0000-C000-000000000046}' : 'SharedWorkspaceFolder',
	'{000C037E-0000-0000-C000-000000000046}' : 'SharedWorkspaceFolders',
	'{000C037F-0000-0000-C000-000000000046}' : 'SharedWorkspaceLink',
	'{000C0380-0000-0000-C000-000000000046}' : 'SharedWorkspaceLinks',
	'{000C0381-0000-0000-C000-000000000046}' : 'SharedWorkspaceMember',
	'{000C0382-0000-0000-C000-000000000046}' : 'SharedWorkspaceMembers',
	'{000C0385-0000-0000-C000-000000000046}' : 'SharedWorkspace',
	'{000C0386-0000-0000-C000-000000000046}' : 'Sync',
	'{000C0387-0000-0000-C000-000000000046}' : 'DocumentLibraryVersion',
	'{000C0388-0000-0000-C000-000000000046}' : 'DocumentLibraryVersions',
	'{000C0389-0000-0000-C000-000000000046}' : 'MsoDebugOptions_UTManager',
	'{000C038A-0000-0000-C000-000000000046}' : 'MsoDebugOptions_UTs',
	'{000C038B-0000-0000-C000-000000000046}' : 'MsoDebugOptions_UT',
	'{000C038C-0000-0000-C000-000000000046}' : 'MsoDebugOptions_UTRunResult',
	'{000C038E-0000-0000-C000-000000000046}' : 'MetaProperties',
	'{000C038F-0000-0000-C000-000000000046}' : 'MetaProperty',
	'{000C0390-0000-0000-C000-000000000046}' : 'ServerPolicy',
	'{000C0391-0000-0000-C000-000000000046}' : 'PolicyItem',
	'{000C0392-0000-0000-C000-000000000046}' : 'DocumentInspectors',
	'{000C0393-0000-0000-C000-000000000046}' : 'DocumentInspector',
	'{000C0395-0000-0000-C000-000000000046}' : 'IRibbonControl',
	'{00194002-D9C3-11D3-8D59-0050048384E3}' : 'ILicAgent',
	'{000C0397-0000-0000-C000-000000000046}' : 'TextRange2',
	'{000C0398-0000-0000-C000-000000000046}' : 'TextFrame2',
	'{000C0399-0000-0000-C000-000000000046}' : 'ParagraphFormat2',
	'{000C039A-0000-0000-C000-000000000046}' : 'Font2',
	'{000C03A0-0000-0000-C000-000000000046}' : 'OfficeTheme',
	'{000C03A1-0000-0000-C000-000000000046}' : 'ThemeColor',
	'{000C03A2-0000-0000-C000-000000000046}' : 'ThemeColorScheme',
	'{000C03A3-0000-0000-C000-000000000046}' : 'ThemeFont',
	'{000C03A4-0000-0000-C000-000000000046}' : 'ThemeFonts',
	'{000C03A5-0000-0000-C000-000000000046}' : 'ThemeFontScheme',
	'{000C03A6-0000-0000-C000-000000000046}' : 'ThemeEffectScheme',
	'{000C03A7-0000-0000-C000-000000000046}' : 'IRibbonUI',
	'{919AA22C-B9AD-11D3-8D59-0050048384E3}' : 'ILicValidator',
	'{000C03B2-0000-0000-C000-000000000046}' : 'TextColumn2',
	'{000C03B9-0000-0000-C000-000000000046}' : 'BulletFormat2',
	'{000C03BA-0000-0000-C000-000000000046}' : 'TabStops2',
	'{000C03BB-0000-0000-C000-000000000046}' : 'TabStop2',
	'{000C03BC-0000-0000-C000-000000000046}' : 'SoftEdgeFormat',
	'{000C03BD-0000-0000-C000-000000000046}' : 'GlowFormat',
	'{000C03BE-0000-0000-C000-000000000046}' : 'ReflectionFormat',
	'{000C03BF-0000-0000-C000-000000000046}' : 'GradientStop',
	'{000C03C0-0000-0000-C000-000000000046}' : 'GradientStops',
	'{000C0300-0000-0000-C000-000000000046}' : '_IMsoDispObj',
	'{4291224C-DEFE-485B-8E69-6CF8AA85CB76}' : 'IAssistance',
	'{000C03C3-0000-0000-C000-000000000046}' : 'RulerLevel2',
	'{000C03C4-0000-0000-C000-000000000046}' : 'IBlogExtensibility',
	'{000C03C5-0000-0000-C000-000000000046}' : 'IBlogPictureExtensibility',
	'{000C0301-0000-0000-C000-000000000046}' : '_IMsoOleAccDispObj',
	'{000C0302-0000-0000-C000-000000000046}' : '_CommandBars',
	'{000CDB03-0000-0000-C000-000000000046}' : 'CustomXMLNodes',
	'{ABFA087C-F703-4D53-946E-37FF82B2C994}' : 'IMsoAxisTitle',
	'{000CDB04-0000-0000-C000-000000000046}' : 'CustomXMLNode',
	'{000CD706-0000-0000-C000-000000000046}' : 'IDocumentInspector',
	'{000C1709-0000-0000-C000-000000000046}' : 'IMsoChart',
	'{000CDB0A-0000-0000-C000-000000000046}' : 'ICustomXMLPartsEvents',
	'{000CD809-0000-0000-C000-000000000046}' : 'EncryptionProvider',
	'{000C0410-0000-0000-C000-000000000046}' : 'SignatureSet',
	'{000C0411-0000-0000-C000-000000000046}' : 'Signature',
	'{000CDB0E-0000-0000-C000-000000000046}' : 'CustomXMLValidationError',
	'{000C170F-0000-0000-C000-000000000046}' : 'IMsoChartTitle',
	'{000CDB10-0000-0000-C000-000000000046}' : 'CustomXMLPrefixMapping',
	'{A98639A1-CB0C-4A5C-A511-96547F752ACD}' : 'IMsoHyperlinks',
	'{000C1730-0000-0000-C000-000000000046}' : 'IMsoChartFormat',
	'{000C1711-0000-0000-C000-000000000046}' : 'IMsoDataTable',
	'{000C1712-0000-0000-C000-000000000046}' : 'Axes',
	'{000C0313-0000-0000-C000-000000000046}' : 'ConnectorFormat',
	'{000C0314-0000-0000-C000-000000000046}' : 'FillFormat',
	'{000C1715-0000-0000-C000-000000000046}' : 'IMsoWalls',
	'{000C1731-0000-0000-C000-000000000046}' : 'IMsoCharacters',
	'{000C1716-0000-0000-C000-000000000046}' : 'IMsoFloor',
	'{000C03C1-0000-0000-C000-000000000046}' : 'Ruler2',
	'{000C1717-0000-0000-C000-000000000046}' : 'IMsoBorder',
	'{000C03C2-0000-0000-C000-000000000046}' : 'RulerLevels2',
	'{000C1718-0000-0000-C000-000000000046}' : 'ChartFont',
	'{000C171B-0000-0000-C000-000000000046}' : 'IMsoInterior',
	'{000C1710-0000-0000-C000-000000000046}' : 'IMsoLegend',
	'{000C171C-0000-0000-C000-000000000046}' : 'ChartFillFormat',
	'{55F88890-7708-11D1-ACEB-006008961DA5}' : 'ICommandBarButtonEvents',
	'{6EA00553-9439-4D5A-B1E6-DC15A54DA8B2}' : 'IMsoDisplayUnitLabel',
	'{55F88892-7708-11D1-ACEB-006008961DA5}' : 'ICommandBarsEvents',
	'{55F88896-7708-11D1-ACEB-006008961DA5}' : 'ICommandBarComboBoxEvents',
	'{000CD900-0000-0000-C000-000000000046}' : 'WorkflowTask',
	'{000C1724-0000-0000-C000-000000000046}' : 'IMsoPlotArea',
	'{000CD6A1-0000-0000-C000-000000000046}' : 'SignatureSetup',
	'{000CD6A2-0000-0000-C000-000000000046}' : 'SignatureInfo',
	'{000CD6A3-0000-0000-C000-000000000046}' : 'SignatureProvider',
	'{000CDB00-0000-0000-C000-000000000046}' : 'CustomXMLPrefixMappings',
	'{000C1726-0000-0000-C000-000000000046}' : 'IMsoTickLabels',
	'{4CAC6328-B9B0-11D3-8D59-0050048384E3}' : 'ILicWizExternal',
	'{000672AC-0000-0000-C000-000000000046}' : 'IMsoEnvelopeVB',
	'{000C1728-0000-0000-C000-000000000046}' : 'IMsoChartArea',
	'{000CD901-0000-0000-C000-000000000046}' : 'WorkflowTasks',
	'{000C1713-0000-0000-C000-000000000046}' : 'IMsoAxis',
	'{000CDB01-0000-0000-C000-000000000046}' : 'CustomXMLSchema',
	'{000C172C-0001-0000-C000-000000000046}' : 'IMsoDropLines',
	'{000CD902-0000-0000-C000-000000000046}' : 'WorkflowTemplate',
	'{8A64A872-FC6B-4D4A-926E-3A3689562C1C}' : 'CustomTaskPaneEvents',
	'{000C172E-0000-0000-C000-000000000046}' : 'IMsoHiLoLines',
	'{000C1714-0000-0000-C000-000000000046}' : 'IMsoCorners',
	'{618736E0-3C3D-11CF-810C-00AA00389B71}' : 'IAccessible',
	'{000CDB02-0000-0000-C000-000000000046}' : '_CustomXMLSchemaCollection',
	'{000C0330-0000-0000-C000-000000000046}' : 'BalloonLabel',
	'{000C0331-0000-0000-C000-000000000046}' : 'FoundFiles',
	'{000C0332-0000-0000-C000-000000000046}' : 'FileSearch',
	'{000CDB06-0000-0000-C000-000000000046}' : 'ICustomXMLPartEvents',
	'{000C0333-0000-0000-C000-000000000046}' : 'PropertyTest',
	'{000C0334-0000-0000-C000-000000000046}' : 'PropertyTests',
	'{000CD100-0000-0000-C000-000000000046}' : 'WebComponent',
	'{000CD101-0000-0000-C000-000000000046}' : 'WebComponentWindowExternal',
	'{000CD102-0000-0000-C000-000000000046}' : 'WebComponentFormat',
	'{000CD903-0000-0000-C000-000000000046}' : 'WorkflowTemplates',
	'{000C0304-0000-0000-C000-000000000046}' : 'CommandBar',
	'{000CDB05-0000-0000-C000-000000000046}' : '_CustomXMLPart',
	'{000C0306-0000-0000-C000-000000000046}' : 'CommandBarControls',
	'{000C0308-0000-0000-C000-000000000046}' : 'CommandBarControl',
	'{000CDB09-0000-0000-C000-000000000046}' : '_CustomXMLParts',
	'{000C030A-0000-0000-C000-000000000046}' : 'CommandBarPopup',
	'{000C030C-0000-0000-C000-000000000046}' : '_CommandBarComboBox',
	'{000C030D-0000-0000-C000-000000000046}' : '_CommandBarActiveX',
	'{000C030E-0000-0000-C000-000000000046}' : '_CommandBarButton',
	'{000CDB0F-0000-0000-C000-000000000046}' : 'CustomXMLValidationErrors',
	'{000C0310-0000-0000-C000-000000000046}' : 'Adjustments',
	'{000C0311-0000-0000-C000-000000000046}' : 'CalloutFormat',
	'{000C0312-0000-0000-C000-000000000046}' : 'ColorFormat',
	'{000C0913-0000-0000-C000-000000000046}' : 'WebPageFont',
	'{000C0914-0000-0000-C000-000000000046}' : 'WebPageFonts',
	'{000C0315-0000-0000-C000-000000000046}' : 'FreeformBuilder',
	'{000C0316-0000-0000-C000-000000000046}' : 'GroupShapes',
	'{000C0317-0000-0000-C000-000000000046}' : 'LineFormat',
	'{000C0318-0000-0000-C000-000000000046}' : 'ShapeNode',
	'{000C0319-0000-0000-C000-000000000046}' : 'ShapeNodes',
	'{000C031A-0000-0000-C000-000000000046}' : 'PictureFormat',
	'{000C031B-0000-0000-C000-000000000046}' : 'ShadowFormat',
	'{000C031C-0000-0000-C000-000000000046}' : 'Shape',
	'{000C031D-0000-0000-C000-000000000046}' : 'ShapeRange',
	'{000C031E-0000-0000-C000-000000000046}' : 'Shapes',
	'{000C031F-0000-0000-C000-000000000046}' : 'TextEffectFormat',
	'{000C0320-0000-0000-C000-000000000046}' : 'TextFrame',
	'{000C0321-0000-0000-C000-000000000046}' : 'ThreeDFormat',
	'{000C0322-0000-0000-C000-000000000046}' : 'Assistant',
	'{000C1723-0000-0000-C000-000000000046}' : 'IMsoLeaderLines',
	'{000C0324-0000-0000-C000-000000000046}' : 'Balloon',
	'{000C1725-0000-0000-C000-000000000046}' : 'GridLines',
	'{000C0326-0000-0000-C000-000000000046}' : 'BalloonCheckboxes',
	'{000C1727-0000-0000-C000-000000000046}' : 'IMsoChartGroup',
	'{000C0328-0000-0000-C000-000000000046}' : 'BalloonCheckbox',
	'{000C1729-0000-0000-C000-000000000046}' : 'IMsoSeriesLines',
	'{000C172A-0000-0000-C000-000000000046}' : 'IMsoUpBars',
	'{000C172B-0000-0000-C000-000000000046}' : 'ChartGroups',
	'{000C172D-0000-0000-C000-000000000046}' : 'IMsoDownBars',
	'{000C032E-0000-0000-C000-000000000046}' : 'BalloonLabels',
	'{000C172F-0000-0000-C000-000000000046}' : 'IMsoChartData',
	'{000C1530-0000-0000-C000-000000000046}' : 'OfficeDataSourceObject',
	'{000C1531-0000-0000-C000-000000000046}' : 'ODSOColumn',
	'{000C1532-0000-0000-C000-000000000046}' : 'ODSOColumns',
	'{000C1533-0000-0000-C000-000000000046}' : 'ODSOFilter',
	'{000C1534-0000-0000-C000-000000000046}' : 'ODSOFilters',
	'{000C0936-0000-0000-C000-000000000046}' : 'NewFile',
	'{000C0337-0000-0000-C000-000000000046}' : 'IFind',
	'{000C0338-0000-0000-C000-000000000046}' : 'IFoundFiles',
	'{000C0339-0000-0000-C000-000000000046}' : 'COMAddIns',
	'{000C033A-0000-0000-C000-000000000046}' : 'COMAddIn',
	'{000C033B-0000-0000-C000-000000000046}' : '_CustomTaskPane',
	'{000C033D-0000-0000-C000-000000000046}' : 'ICTPFactory',
	'{000C033E-0000-0000-C000-000000000046}' : 'ICustomTaskPaneConsumer',
}


NamesToIIDMap = {
	'FileDialogFilters' : '{000C0365-0000-0000-C000-000000000046}',
	'TextRange2' : '{000C0397-0000-0000-C000-000000000046}',
	'ODSOFilter' : '{000C1533-0000-0000-C000-000000000046}',
	'MsoDebugOptions_UT' : '{000C038B-0000-0000-C000-000000000046}',
	'TextColumn2' : '{000C03B2-0000-0000-C000-000000000046}',
	'EncryptionProvider' : '{000CD809-0000-0000-C000-000000000046}',
	'CanvasShapes' : '{000C0371-0000-0000-C000-000000000046}',
	'SharedWorkspaceFolder' : '{000C037D-0000-0000-C000-000000000046}',
	'ThemeEffectScheme' : '{000C03A6-0000-0000-C000-000000000046}',
	'WebPageFont' : '{000C0913-0000-0000-C000-000000000046}',
	'IMsoChartData' : '{000C172F-0000-0000-C000-000000000046}',
	'IMsoDataLabels' : '{000C171F-0000-0000-C000-000000000046}',
	'PropertyTest' : '{000C0333-0000-0000-C000-000000000046}',
	'BalloonCheckbox' : '{000C0328-0000-0000-C000-000000000046}',
	'DocumentLibraryVersion' : '{000C0387-0000-0000-C000-000000000046}',
	'ThreeDFormat' : '{000C0321-0000-0000-C000-000000000046}',
	'BalloonCheckboxes' : '{000C0326-0000-0000-C000-000000000046}',
	'WebComponentFormat' : '{000CD102-0000-0000-C000-000000000046}',
	'GradientStops' : '{000C03C0-0000-0000-C000-000000000046}',
	'ICommandBarButtonEvents' : '{55F88890-7708-11D1-ACEB-006008961DA5}',
	'FileSearch' : '{000C0332-0000-0000-C000-000000000046}',
	'IMsoLegendKey' : '{000C171E-0000-0000-C000-000000000046}',
	'IMsoLeaderLines' : '{000C1723-0000-0000-C000-000000000046}',
	'IRibbonUI' : '{000C03A7-0000-0000-C000-000000000046}',
	'FileDialogFilter' : '{000C0364-0000-0000-C000-000000000046}',
	'MsoDebugOptions_UTRunResult' : '{000C038C-0000-0000-C000-000000000046}',
	'IFoundFiles' : '{000C0338-0000-0000-C000-000000000046}',
	'LanguageSettings' : '{000C0353-0000-0000-C000-000000000046}',
	'WebPageFonts' : '{000C0914-0000-0000-C000-000000000046}',
	'OfficeTheme' : '{000C03A0-0000-0000-C000-000000000046}',
	'ODSOFilters' : '{000C1534-0000-0000-C000-000000000046}',
	'_CommandBarButton' : '{000C030E-0000-0000-C000-000000000046}',
	'IMsoDataLabel' : '{000C1720-0000-0000-C000-000000000046}',
	'IMsoEnvelopeVB' : '{000672AC-0000-0000-C000-000000000046}',
	'ODSOColumn' : '{000C1531-0000-0000-C000-000000000046}',
	'SignatureProvider' : '{000CD6A3-0000-0000-C000-000000000046}',
	'SignatureInfo' : '{000CD6A2-0000-0000-C000-000000000046}',
	'IMsoChartArea' : '{000C1728-0000-0000-C000-000000000046}',
	'SoftEdgeFormat' : '{000C03BC-0000-0000-C000-000000000046}',
	'CustomXMLNodes' : '{000CDB03-0000-0000-C000-000000000046}',
	'IMsoTrendline' : '{000C170E-0000-0000-C000-000000000046}',
	'PictureFormat' : '{000C031A-0000-0000-C000-000000000046}',
	'ThemeFonts' : '{000C03A4-0000-0000-C000-000000000046}',
	'IMsoChart' : '{000C1709-0000-0000-C000-000000000046}',
	'ODSOColumns' : '{000C1532-0000-0000-C000-000000000046}',
	'PolicyItem' : '{000C0391-0000-0000-C000-000000000046}',
	'IMsoCorners' : '{000C1714-0000-0000-C000-000000000046}',
	'FileDialogSelectedItems' : '{000C0363-0000-0000-C000-000000000046}',
	'SharedWorkspaceFile' : '{000C037B-0000-0000-C000-000000000046}',
	'IMsoAxis' : '{000C1713-0000-0000-C000-000000000046}',
	'HTMLProject' : '{000C0356-0000-0000-C000-000000000046}',
	'ShapeNodes' : '{000C0319-0000-0000-C000-000000000046}',
	'ThemeFont' : '{000C03A3-0000-0000-C000-000000000046}',
	'ScopeFolder' : '{000C0368-0000-0000-C000-000000000046}',
	'DocumentLibraryVersions' : '{000C0388-0000-0000-C000-000000000046}',
	'TabStop2' : '{000C03BB-0000-0000-C000-000000000046}',
	'IMsoChartTitle' : '{000C170F-0000-0000-C000-000000000046}',
	'_CustomXMLPartsEvents' : '{000CDB0B-0000-0000-C000-000000000046}',
	'Script' : '{000C0341-0000-0000-C000-000000000046}',
	'BalloonLabels' : '{000C032E-0000-0000-C000-000000000046}',
	'Balloon' : '{000C0324-0000-0000-C000-000000000046}',
	'Trendlines' : '{000C1722-0000-0000-C000-000000000046}',
	'IMsoLegend' : '{000C1710-0000-0000-C000-000000000046}',
	'IMsoHyperlinks' : '{A98639A1-CB0C-4A5C-A511-96547F752ACD}',
	'SearchFolders' : '{000C036A-0000-0000-C000-000000000046}',
	'IMsoEServicesDialog' : '{000C0372-0000-0000-C000-000000000046}',
	'SignatureSetup' : '{000CD6A1-0000-0000-C000-000000000046}',
	'Permission' : '{000C0376-0000-0000-C000-000000000046}',
	'IAccessible' : '{618736E0-3C3D-11CF-810C-00AA00389B71}',
	'IMsoWalls' : '{000C1715-0000-0000-C000-000000000046}',
	'SharedWorkspaceFolders' : '{000C037E-0000-0000-C000-000000000046}',
	'ShapeRange' : '{000C031D-0000-0000-C000-000000000046}',
	'ICustomXMLPartsEvents' : '{000CDB0A-0000-0000-C000-000000000046}',
	'MsoDebugOptions_UTs' : '{000C038A-0000-0000-C000-000000000046}',
	'_CommandBarActiveX' : '{000C030D-0000-0000-C000-000000000046}',
	'FillFormat' : '{000C0314-0000-0000-C000-000000000046}',
	'CustomXMLValidationErrors' : '{000CDB0F-0000-0000-C000-000000000046}',
	'IBlogExtensibility' : '{000C03C4-0000-0000-C000-000000000046}',
	'SeriesCollection' : '{000C170A-0000-0000-C000-000000000046}',
	'FoundFiles' : '{000C0331-0000-0000-C000-000000000046}',
	'COMAddIns' : '{000C0339-0000-0000-C000-000000000046}',
	'WebComponentWindowExternal' : '{000CD101-0000-0000-C000-000000000046}',
	'HTMLProjectItems' : '{000C0357-0000-0000-C000-000000000046}',
	'Adjustments' : '{000C0310-0000-0000-C000-000000000046}',
	'IMsoSeries' : '{000C170B-0000-0000-C000-000000000046}',
	'TextEffectFormat' : '{000C031F-0000-0000-C000-000000000046}',
	'SmartDocument' : '{000C0377-0000-0000-C000-000000000046}',
	'CustomXMLNode' : '{000CDB04-0000-0000-C000-000000000046}',
	'_CustomXMLPart' : '{000CDB05-0000-0000-C000-000000000046}',
	'WebComponent' : '{000CD100-0000-0000-C000-000000000046}',
	'SharedWorkspaceMember' : '{000C0381-0000-0000-C000-000000000046}',
	'ChartFont' : '{000C1718-0000-0000-C000-000000000046}',
	'IMsoInterior' : '{000C171B-0000-0000-C000-000000000046}',
	'WebComponentProperties' : '{000C0373-0000-0000-C000-000000000046}',
	'IMsoCharacters' : '{000C1731-0000-0000-C000-000000000046}',
	'Signature' : '{000C0411-0000-0000-C000-000000000046}',
	'CalloutFormat' : '{000C0311-0000-0000-C000-000000000046}',
	'GradientStop' : '{000C03BF-0000-0000-C000-000000000046}',
	'FreeformBuilder' : '{000C0315-0000-0000-C000-000000000046}',
	'BulletFormat2' : '{000C03B9-0000-0000-C000-000000000046}',
	'_CustomXMLParts' : '{000CDB09-0000-0000-C000-000000000046}',
	'OfficeDataSourceObject' : '{000C1530-0000-0000-C000-000000000046}',
	'Shapes' : '{000C031E-0000-0000-C000-000000000046}',
	'IMsoBorder' : '{000C1717-0000-0000-C000-000000000046}',
	'SearchScopes' : '{000C0366-0000-0000-C000-000000000046}',
	'IMsoChartFormat' : '{000C1730-0000-0000-C000-000000000046}',
	'IMsoSeriesLines' : '{000C1729-0000-0000-C000-000000000046}',
	'IFind' : '{000C0337-0000-0000-C000-000000000046}',
	'SharedWorkspaceMembers' : '{000C0382-0000-0000-C000-000000000046}',
	'Shape' : '{000C031C-0000-0000-C000-000000000046}',
	'SharedWorkspaceTasks' : '{000C037A-0000-0000-C000-000000000046}',
	'LineFormat' : '{000C0317-0000-0000-C000-000000000046}',
	'_IMsoDispObj' : '{000C0300-0000-0000-C000-000000000046}',
	'ChartPoint' : '{000C170C-0000-0000-C000-000000000046}',
	'SharedWorkspaceLinks' : '{000C0380-0000-0000-C000-000000000046}',
	'ThemeColor' : '{000C03A1-0000-0000-C000-000000000046}',
	'RulerLevels2' : '{000C03C2-0000-0000-C000-000000000046}',
	'MsoDebugOptions_UTManager' : '{000C0389-0000-0000-C000-000000000046}',
	'Font2' : '{000C039A-0000-0000-C000-000000000046}',
	'Axes' : '{000C1712-0000-0000-C000-000000000046}',
	'SharedWorkspaceFiles' : '{000C037C-0000-0000-C000-000000000046}',
	'MsoDebugOptions' : '{000C035A-0000-0000-C000-000000000046}',
	'ServerPolicy' : '{000C0390-0000-0000-C000-000000000046}',
	'IMsoPlotArea' : '{000C1724-0000-0000-C000-000000000046}',
	'COMAddIn' : '{000C033A-0000-0000-C000-000000000046}',
	'FileDialog' : '{000C0362-0000-0000-C000-000000000046}',
	'TextFrame2' : '{000C0398-0000-0000-C000-000000000046}',
	'ColorFormat' : '{000C0312-0000-0000-C000-000000000046}',
	'ChartGroups' : '{000C172B-0000-0000-C000-000000000046}',
	'IMsoDownBars' : '{000C172D-0000-0000-C000-000000000046}',
	'DiagramNode' : '{000C0370-0000-0000-C000-000000000046}',
	'IMsoHiLoLines' : '{000C172E-0000-0000-C000-000000000046}',
	'TextFrame' : '{000C0320-0000-0000-C000-000000000046}',
	'Assistant' : '{000C0322-0000-0000-C000-000000000046}',
	'IMsoChartGroup' : '{000C1727-0000-0000-C000-000000000046}',
	'IMsoDataTable' : '{000C1711-0000-0000-C000-000000000046}',
	'ReflectionFormat' : '{000C03BE-0000-0000-C000-000000000046}',
	'ScopeFolders' : '{000C0369-0000-0000-C000-000000000046}',
	'CommandBar' : '{000C0304-0000-0000-C000-000000000046}',
	'DiagramNodes' : '{000C036E-0000-0000-C000-000000000046}',
	'ConnectorFormat' : '{000C0313-0000-0000-C000-000000000046}',
	'ICustomTaskPaneConsumer' : '{000C033E-0000-0000-C000-000000000046}',
	'BalloonLabel' : '{000C0330-0000-0000-C000-000000000046}',
	'MetaProperties' : '{000C038E-0000-0000-C000-000000000046}',
	'NewFile' : '{000C0936-0000-0000-C000-000000000046}',
	'_CommandBarComboBox' : '{000C030C-0000-0000-C000-000000000046}',
	'WorkflowTask' : '{000CD900-0000-0000-C000-000000000046}',
	'WorkflowTemplates' : '{000CD903-0000-0000-C000-000000000046}',
	'IAssistance' : '{4291224C-DEFE-485B-8E69-6CF8AA85CB76}',
	'ICommandBarsEvents' : '{55F88892-7708-11D1-ACEB-006008961DA5}',
	'DiagramNodeChildren' : '{000C036F-0000-0000-C000-000000000046}',
	'AnswerWizardFiles' : '{000C0361-0000-0000-C000-000000000046}',
	'Points' : '{000C170D-0000-0000-C000-000000000046}',
	'CustomXMLSchema' : '{000CDB01-0000-0000-C000-000000000046}',
	'IMsoDiagram' : '{000C036D-0000-0000-C000-000000000046}',
	'CustomTaskPaneEvents' : '{8A64A872-FC6B-4D4A-926E-3A3689562C1C}',
	'IMsoErrorBars' : '{000C1721-0000-0000-C000-000000000046}',
	'IMsoAxisTitle' : '{ABFA087C-F703-4D53-946E-37FF82B2C994}',
	'ParagraphFormat2' : '{000C0399-0000-0000-C000-000000000046}',
	'_CommandBarsEvents' : '{000C0352-0000-0000-C000-000000000046}',
	'SharedWorkspaceTask' : '{000C0379-0000-0000-C000-000000000046}',
	'ShapeNode' : '{000C0318-0000-0000-C000-000000000046}',
	'DocumentProperty' : '{2DF8D04E-5BFA-101B-BDE5-00AA0044DE52}',
	'IDocumentInspector' : '{000CD706-0000-0000-C000-000000000046}',
	'IRibbonControl' : '{000C0395-0000-0000-C000-000000000046}',
	'LegendEntries' : '{000C1719-0000-0000-C000-000000000046}',
	'ILicWizExternal' : '{4CAC6328-B9B0-11D3-8D59-0050048384E3}',
	'_CustomXMLPartEvents' : '{000CDB07-0000-0000-C000-000000000046}',
	'CustomXMLPrefixMapping' : '{000CDB10-0000-0000-C000-000000000046}',
	'DocumentInspectors' : '{000C0392-0000-0000-C000-000000000046}',
	'ChartFillFormat' : '{000C171C-0000-0000-C000-000000000046}',
	'_CustomTaskPane' : '{000C033B-0000-0000-C000-000000000046}',
	'SignatureSet' : '{000C0410-0000-0000-C000-000000000046}',
	'SearchScope' : '{000C0367-0000-0000-C000-000000000046}',
	'MetaProperty' : '{000C038F-0000-0000-C000-000000000046}',
	'IMsoDropLines' : '{000C172C-0001-0000-C000-000000000046}',
	'ChartColorFormat' : '{000C171D-0000-0000-C000-000000000046}',
	'_CustomXMLSchemaCollection' : '{000CDB02-0000-0000-C000-000000000046}',
	'Sync' : '{000C0386-0000-0000-C000-000000000046}',
	'CustomXMLValidationError' : '{000CDB0E-0000-0000-C000-000000000046}',
	'UserPermission' : '{000C0375-0000-0000-C000-000000000046}',
	'ILicValidator' : '{919AA22C-B9AD-11D3-8D59-0050048384E3}',
	'CommandBarControls' : '{000C0306-0000-0000-C000-000000000046}',
	'IRibbonExtensibility' : '{000C0396-0000-0000-C000-000000000046}',
	'WorkflowTasks' : '{000CD901-0000-0000-C000-000000000046}',
	'GridLines' : '{000C1725-0000-0000-C000-000000000046}',
	'TabStops2' : '{000C03BA-0000-0000-C000-000000000046}',
	'AnswerWizard' : '{000C0360-0000-0000-C000-000000000046}',
	'ICTPFactory' : '{000C033D-0000-0000-C000-000000000046}',
	'CustomXMLPrefixMappings' : '{000CDB00-0000-0000-C000-000000000046}',
	'_IMsoOleAccDispObj' : '{000C0301-0000-0000-C000-000000000046}',
	'Ruler2' : '{000C03C1-0000-0000-C000-000000000046}',
	'SharedWorkspaceLink' : '{000C037F-0000-0000-C000-000000000046}',
	'ThemeFontScheme' : '{000C03A5-0000-0000-C000-000000000046}',
	'_CommandBarButtonEvents' : '{000C0351-0000-0000-C000-000000000046}',
	'_CommandBarComboBoxEvents' : '{000C0354-0000-0000-C000-000000000046}',
	'Scripts' : '{000C0340-0000-0000-C000-000000000046}',
	'IMsoDisplayUnitLabel' : '{6EA00553-9439-4D5A-B1E6-DC15A54DA8B2}',
	'FileTypes' : '{000C036C-0000-0000-C000-000000000046}',
	'_CommandBars' : '{000C0302-0000-0000-C000-000000000046}',
	'ILicAgent' : '{00194002-D9C3-11D3-8D59-0050048384E3}',
	'IMsoTickLabels' : '{000C1726-0000-0000-C000-000000000046}',
	'LegendEntry' : '{000C171A-0000-0000-C000-000000000046}',
	'DocumentInspector' : '{000C0393-0000-0000-C000-000000000046}',
	'IBlogPictureExtensibility' : '{000C03C5-0000-0000-C000-000000000046}',
	'DocumentProperties' : '{2DF8D04D-5BFA-101B-BDE5-00AA0044DE52}',
	'ThemeColorScheme' : '{000C03A2-0000-0000-C000-000000000046}',
	'IMsoDispCagNotifySink' : '{000C0359-0000-0000-C000-000000000046}',
	'ShadowFormat' : '{000C031B-0000-0000-C000-000000000046}',
	'PropertyTests' : '{000C0334-0000-0000-C000-000000000046}',
	'_CustomTaskPaneEvents' : '{000C033C-0000-0000-C000-000000000046}',
	'HTMLProjectItem' : '{000C0358-0000-0000-C000-000000000046}',
	'RulerLevel2' : '{000C03C3-0000-0000-C000-000000000046}',
	'IMsoEnvelopeVBEvents' : '{000672AD-0000-0000-C000-000000000046}',
	'ICustomXMLPartEvents' : '{000CDB06-0000-0000-C000-000000000046}',
	'CommandBarControl' : '{000C0308-0000-0000-C000-000000000046}',
	'GlowFormat' : '{000C03BD-0000-0000-C000-000000000046}',
	'WorkflowTemplate' : '{000CD902-0000-0000-C000-000000000046}',
	'GroupShapes' : '{000C0316-0000-0000-C000-000000000046}',
	'SharedWorkspace' : '{000C0385-0000-0000-C000-000000000046}',
	'ICommandBarComboBoxEvents' : '{55F88896-7708-11D1-ACEB-006008961DA5}',
	'IMsoUpBars' : '{000C172A-0000-0000-C000-000000000046}',
	'CommandBarPopup' : '{000C030A-0000-0000-C000-000000000046}',
	'IMsoFloor' : '{000C1716-0000-0000-C000-000000000046}',
}

win32com.client.constants.__dicts__.append(constants.__dict__)

