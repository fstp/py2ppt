# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 2.7.8 (default, Jun 30 2014, 16:03:49) [MSC v.1500 32 bit (Intel)]
# From type library 'MSPPT.OLB'
# On Sun Oct 04 23:15:13 2015
'Microsoft PowerPoint 12.0 Object Library'
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

CLSID = IID('{91493440-5A91-11CF-8700-00AA0060263B}')
MajorVersion = 2
MinorVersion = 9
LibraryFlags = 8
LCID = 0x0

class constants:
	msoAnimAccumulateAlways       =2          # from enum MsoAnimAccumulate
	msoAnimAccumulateNone         =1          # from enum MsoAnimAccumulate
	msoAnimAdditiveAddBase        =1          # from enum MsoAnimAdditive
	msoAnimAdditiveAddSum         =2          # from enum MsoAnimAdditive
	msoAnimAfterEffectDim         =1          # from enum MsoAnimAfterEffect
	msoAnimAfterEffectHide        =2          # from enum MsoAnimAfterEffect
	msoAnimAfterEffectHideOnNextClick=3          # from enum MsoAnimAfterEffect
	msoAnimAfterEffectMixed       =-1         # from enum MsoAnimAfterEffect
	msoAnimAfterEffectNone        =0          # from enum MsoAnimAfterEffect
	msoAnimCommandTypeCall        =1          # from enum MsoAnimCommandType
	msoAnimCommandTypeEvent       =0          # from enum MsoAnimCommandType
	msoAnimCommandTypeVerb        =2          # from enum MsoAnimCommandType
	msoAnimDirectionAcross        =18         # from enum MsoAnimDirection
	msoAnimDirectionBottom        =11         # from enum MsoAnimDirection
	msoAnimDirectionBottomLeft    =15         # from enum MsoAnimDirection
	msoAnimDirectionBottomRight   =14         # from enum MsoAnimDirection
	msoAnimDirectionCenter        =28         # from enum MsoAnimDirection
	msoAnimDirectionClockwise     =21         # from enum MsoAnimDirection
	msoAnimDirectionCounterclockwise=22         # from enum MsoAnimDirection
	msoAnimDirectionCycleClockwise=43         # from enum MsoAnimDirection
	msoAnimDirectionCycleCounterclockwise=44         # from enum MsoAnimDirection
	msoAnimDirectionDown          =3          # from enum MsoAnimDirection
	msoAnimDirectionDownLeft      =9          # from enum MsoAnimDirection
	msoAnimDirectionDownRight     =8          # from enum MsoAnimDirection
	msoAnimDirectionFontAllCaps   =40         # from enum MsoAnimDirection
	msoAnimDirectionFontBold      =35         # from enum MsoAnimDirection
	msoAnimDirectionFontItalic    =36         # from enum MsoAnimDirection
	msoAnimDirectionFontShadow    =39         # from enum MsoAnimDirection
	msoAnimDirectionFontStrikethrough=38         # from enum MsoAnimDirection
	msoAnimDirectionFontUnderline =37         # from enum MsoAnimDirection
	msoAnimDirectionGradual       =42         # from enum MsoAnimDirection
	msoAnimDirectionHorizontal    =16         # from enum MsoAnimDirection
	msoAnimDirectionHorizontalIn  =23         # from enum MsoAnimDirection
	msoAnimDirectionHorizontalOut =24         # from enum MsoAnimDirection
	msoAnimDirectionIn            =19         # from enum MsoAnimDirection
	msoAnimDirectionInBottom      =31         # from enum MsoAnimDirection
	msoAnimDirectionInCenter      =30         # from enum MsoAnimDirection
	msoAnimDirectionInSlightly    =29         # from enum MsoAnimDirection
	msoAnimDirectionInstant       =41         # from enum MsoAnimDirection
	msoAnimDirectionLeft          =4          # from enum MsoAnimDirection
	msoAnimDirectionNone          =0          # from enum MsoAnimDirection
	msoAnimDirectionOrdinalMask   =5          # from enum MsoAnimDirection
	msoAnimDirectionOut           =20         # from enum MsoAnimDirection
	msoAnimDirectionOutBottom     =34         # from enum MsoAnimDirection
	msoAnimDirectionOutCenter     =33         # from enum MsoAnimDirection
	msoAnimDirectionOutSlightly   =32         # from enum MsoAnimDirection
	msoAnimDirectionRight         =2          # from enum MsoAnimDirection
	msoAnimDirectionSlightly      =27         # from enum MsoAnimDirection
	msoAnimDirectionTop           =10         # from enum MsoAnimDirection
	msoAnimDirectionTopLeft       =12         # from enum MsoAnimDirection
	msoAnimDirectionTopRight      =13         # from enum MsoAnimDirection
	msoAnimDirectionUp            =1          # from enum MsoAnimDirection
	msoAnimDirectionUpLeft        =6          # from enum MsoAnimDirection
	msoAnimDirectionUpRight       =7          # from enum MsoAnimDirection
	msoAnimDirectionVertical      =17         # from enum MsoAnimDirection
	msoAnimDirectionVerticalIn    =25         # from enum MsoAnimDirection
	msoAnimDirectionVerticalOut   =26         # from enum MsoAnimDirection
	msoAnimEffectAppear           =1          # from enum MsoAnimEffect
	msoAnimEffectArcUp            =47         # from enum MsoAnimEffect
	msoAnimEffectAscend           =39         # from enum MsoAnimEffect
	msoAnimEffectBlast            =64         # from enum MsoAnimEffect
	msoAnimEffectBlinds           =3          # from enum MsoAnimEffect
	msoAnimEffectBoldFlash        =63         # from enum MsoAnimEffect
	msoAnimEffectBoldReveal       =65         # from enum MsoAnimEffect
	msoAnimEffectBoomerang        =25         # from enum MsoAnimEffect
	msoAnimEffectBounce           =26         # from enum MsoAnimEffect
	msoAnimEffectBox              =4          # from enum MsoAnimEffect
	msoAnimEffectBrushOnColor     =66         # from enum MsoAnimEffect
	msoAnimEffectBrushOnUnderline =67         # from enum MsoAnimEffect
	msoAnimEffectCenterRevolve    =40         # from enum MsoAnimEffect
	msoAnimEffectChangeFillColor  =54         # from enum MsoAnimEffect
	msoAnimEffectChangeFont       =55         # from enum MsoAnimEffect
	msoAnimEffectChangeFontColor  =56         # from enum MsoAnimEffect
	msoAnimEffectChangeFontSize   =57         # from enum MsoAnimEffect
	msoAnimEffectChangeFontStyle  =58         # from enum MsoAnimEffect
	msoAnimEffectChangeLineColor  =60         # from enum MsoAnimEffect
	msoAnimEffectCheckerboard     =5          # from enum MsoAnimEffect
	msoAnimEffectCircle           =6          # from enum MsoAnimEffect
	msoAnimEffectColorBlend       =68         # from enum MsoAnimEffect
	msoAnimEffectColorReveal      =27         # from enum MsoAnimEffect
	msoAnimEffectColorWave        =69         # from enum MsoAnimEffect
	msoAnimEffectComplementaryColor=70         # from enum MsoAnimEffect
	msoAnimEffectComplementaryColor2=71         # from enum MsoAnimEffect
	msoAnimEffectContrastingColor =72         # from enum MsoAnimEffect
	msoAnimEffectCrawl            =7          # from enum MsoAnimEffect
	msoAnimEffectCredits          =28         # from enum MsoAnimEffect
	msoAnimEffectCustom           =0          # from enum MsoAnimEffect
	msoAnimEffectDarken           =73         # from enum MsoAnimEffect
	msoAnimEffectDesaturate       =74         # from enum MsoAnimEffect
	msoAnimEffectDescend          =42         # from enum MsoAnimEffect
	msoAnimEffectDiamond          =8          # from enum MsoAnimEffect
	msoAnimEffectDissolve         =9          # from enum MsoAnimEffect
	msoAnimEffectEaseIn           =29         # from enum MsoAnimEffect
	msoAnimEffectExpand           =50         # from enum MsoAnimEffect
	msoAnimEffectFade             =10         # from enum MsoAnimEffect
	msoAnimEffectFadedSwivel      =41         # from enum MsoAnimEffect
	msoAnimEffectFadedZoom        =48         # from enum MsoAnimEffect
	msoAnimEffectFlashBulb        =75         # from enum MsoAnimEffect
	msoAnimEffectFlashOnce        =11         # from enum MsoAnimEffect
	msoAnimEffectFlicker          =76         # from enum MsoAnimEffect
	msoAnimEffectFlip             =51         # from enum MsoAnimEffect
	msoAnimEffectFloat            =30         # from enum MsoAnimEffect
	msoAnimEffectFly              =2          # from enum MsoAnimEffect
	msoAnimEffectFold             =53         # from enum MsoAnimEffect
	msoAnimEffectGlide            =49         # from enum MsoAnimEffect
	msoAnimEffectGrowAndTurn      =31         # from enum MsoAnimEffect
	msoAnimEffectGrowShrink       =59         # from enum MsoAnimEffect
	msoAnimEffectGrowWithColor    =77         # from enum MsoAnimEffect
	msoAnimEffectLightSpeed       =32         # from enum MsoAnimEffect
	msoAnimEffectLighten          =78         # from enum MsoAnimEffect
	msoAnimEffectMediaPause       =84         # from enum MsoAnimEffect
	msoAnimEffectMediaPlay        =83         # from enum MsoAnimEffect
	msoAnimEffectMediaStop        =85         # from enum MsoAnimEffect
	msoAnimEffectPath4PointStar   =101        # from enum MsoAnimEffect
	msoAnimEffectPath5PointStar   =90         # from enum MsoAnimEffect
	msoAnimEffectPath6PointStar   =96         # from enum MsoAnimEffect
	msoAnimEffectPath8PointStar   =102        # from enum MsoAnimEffect
	msoAnimEffectPathArcDown      =122        # from enum MsoAnimEffect
	msoAnimEffectPathArcLeft      =136        # from enum MsoAnimEffect
	msoAnimEffectPathArcRight     =143        # from enum MsoAnimEffect
	msoAnimEffectPathArcUp        =129        # from enum MsoAnimEffect
	msoAnimEffectPathBean         =116        # from enum MsoAnimEffect
	msoAnimEffectPathBounceLeft   =126        # from enum MsoAnimEffect
	msoAnimEffectPathBounceRight  =139        # from enum MsoAnimEffect
	msoAnimEffectPathBuzzsaw      =110        # from enum MsoAnimEffect
	msoAnimEffectPathCircle       =86         # from enum MsoAnimEffect
	msoAnimEffectPathCrescentMoon =91         # from enum MsoAnimEffect
	msoAnimEffectPathCurvedSquare =105        # from enum MsoAnimEffect
	msoAnimEffectPathCurvedX      =106        # from enum MsoAnimEffect
	msoAnimEffectPathCurvyLeft    =133        # from enum MsoAnimEffect
	msoAnimEffectPathCurvyRight   =146        # from enum MsoAnimEffect
	msoAnimEffectPathCurvyStar    =108        # from enum MsoAnimEffect
	msoAnimEffectPathDecayingWave =145        # from enum MsoAnimEffect
	msoAnimEffectPathDiagonalDownRight=134        # from enum MsoAnimEffect
	msoAnimEffectPathDiagonalUpRight=141        # from enum MsoAnimEffect
	msoAnimEffectPathDiamond      =88         # from enum MsoAnimEffect
	msoAnimEffectPathDown         =127        # from enum MsoAnimEffect
	msoAnimEffectPathEqualTriangle=98         # from enum MsoAnimEffect
	msoAnimEffectPathFigure8Four  =113        # from enum MsoAnimEffect
	msoAnimEffectPathFootball     =97         # from enum MsoAnimEffect
	msoAnimEffectPathFunnel       =137        # from enum MsoAnimEffect
	msoAnimEffectPathHeart        =94         # from enum MsoAnimEffect
	msoAnimEffectPathHeartbeat    =130        # from enum MsoAnimEffect
	msoAnimEffectPathHexagon      =89         # from enum MsoAnimEffect
	msoAnimEffectPathHorizontalFigure8=111        # from enum MsoAnimEffect
	msoAnimEffectPathInvertedSquare=119        # from enum MsoAnimEffect
	msoAnimEffectPathInvertedTriangle=118        # from enum MsoAnimEffect
	msoAnimEffectPathLeft         =120        # from enum MsoAnimEffect
	msoAnimEffectPathLoopdeLoop   =109        # from enum MsoAnimEffect
	msoAnimEffectPathNeutron      =114        # from enum MsoAnimEffect
	msoAnimEffectPathOctagon      =95         # from enum MsoAnimEffect
	msoAnimEffectPathParallelogram=99         # from enum MsoAnimEffect
	msoAnimEffectPathPeanut       =112        # from enum MsoAnimEffect
	msoAnimEffectPathPentagon     =100        # from enum MsoAnimEffect
	msoAnimEffectPathPlus         =117        # from enum MsoAnimEffect
	msoAnimEffectPathPointyStar   =104        # from enum MsoAnimEffect
	msoAnimEffectPathRight        =149        # from enum MsoAnimEffect
	msoAnimEffectPathRightTriangle=87         # from enum MsoAnimEffect
	msoAnimEffectPathSCurve1      =144        # from enum MsoAnimEffect
	msoAnimEffectPathSCurve2      =124        # from enum MsoAnimEffect
	msoAnimEffectPathSineWave     =125        # from enum MsoAnimEffect
	msoAnimEffectPathSpiralLeft   =140        # from enum MsoAnimEffect
	msoAnimEffectPathSpiralRight  =131        # from enum MsoAnimEffect
	msoAnimEffectPathSpring       =138        # from enum MsoAnimEffect
	msoAnimEffectPathSquare       =92         # from enum MsoAnimEffect
	msoAnimEffectPathStairsDown   =147        # from enum MsoAnimEffect
	msoAnimEffectPathSwoosh       =115        # from enum MsoAnimEffect
	msoAnimEffectPathTeardrop     =103        # from enum MsoAnimEffect
	msoAnimEffectPathTrapezoid    =93         # from enum MsoAnimEffect
	msoAnimEffectPathTurnDown     =135        # from enum MsoAnimEffect
	msoAnimEffectPathTurnRight    =121        # from enum MsoAnimEffect
	msoAnimEffectPathTurnUp       =128        # from enum MsoAnimEffect
	msoAnimEffectPathTurnUpRight  =142        # from enum MsoAnimEffect
	msoAnimEffectPathUp           =148        # from enum MsoAnimEffect
	msoAnimEffectPathVerticalFigure8=107        # from enum MsoAnimEffect
	msoAnimEffectPathWave         =132        # from enum MsoAnimEffect
	msoAnimEffectPathZigzag       =123        # from enum MsoAnimEffect
	msoAnimEffectPeek             =12         # from enum MsoAnimEffect
	msoAnimEffectPinwheel         =33         # from enum MsoAnimEffect
	msoAnimEffectPlus             =13         # from enum MsoAnimEffect
	msoAnimEffectRandomBars       =14         # from enum MsoAnimEffect
	msoAnimEffectRandomEffects    =24         # from enum MsoAnimEffect
	msoAnimEffectRiseUp           =34         # from enum MsoAnimEffect
	msoAnimEffectShimmer          =52         # from enum MsoAnimEffect
	msoAnimEffectSling            =43         # from enum MsoAnimEffect
	msoAnimEffectSpin             =61         # from enum MsoAnimEffect
	msoAnimEffectSpinner          =44         # from enum MsoAnimEffect
	msoAnimEffectSpiral           =15         # from enum MsoAnimEffect
	msoAnimEffectSplit            =16         # from enum MsoAnimEffect
	msoAnimEffectStretch          =17         # from enum MsoAnimEffect
	msoAnimEffectStretchy         =45         # from enum MsoAnimEffect
	msoAnimEffectStrips           =18         # from enum MsoAnimEffect
	msoAnimEffectStyleEmphasis    =79         # from enum MsoAnimEffect
	msoAnimEffectSwish            =35         # from enum MsoAnimEffect
	msoAnimEffectSwivel           =19         # from enum MsoAnimEffect
	msoAnimEffectTeeter           =80         # from enum MsoAnimEffect
	msoAnimEffectThinLine         =36         # from enum MsoAnimEffect
	msoAnimEffectTransparency     =62         # from enum MsoAnimEffect
	msoAnimEffectUnfold           =37         # from enum MsoAnimEffect
	msoAnimEffectVerticalGrow     =81         # from enum MsoAnimEffect
	msoAnimEffectWave             =82         # from enum MsoAnimEffect
	msoAnimEffectWedge            =20         # from enum MsoAnimEffect
	msoAnimEffectWheel            =21         # from enum MsoAnimEffect
	msoAnimEffectWhip             =38         # from enum MsoAnimEffect
	msoAnimEffectWipe             =22         # from enum MsoAnimEffect
	msoAnimEffectZip              =46         # from enum MsoAnimEffect
	msoAnimEffectZoom             =23         # from enum MsoAnimEffect
	msoAnimEffectAfterFreeze      =1          # from enum MsoAnimEffectAfter
	msoAnimEffectAfterHold        =3          # from enum MsoAnimEffectAfter
	msoAnimEffectAfterRemove      =2          # from enum MsoAnimEffectAfter
	msoAnimEffectAfterTransition  =4          # from enum MsoAnimEffectAfter
	msoAnimEffectRestartAlways    =1          # from enum MsoAnimEffectRestart
	msoAnimEffectRestartNever     =3          # from enum MsoAnimEffectRestart
	msoAnimEffectRestartWhenOff   =2          # from enum MsoAnimEffectRestart
	msoAnimFilterEffectSubtypeAcross=9          # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeDown=25         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeDownLeft=14         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeDownRight=16         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeFromBottom=13         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeFromLeft=10         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeFromRight=11         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeFromTop=12         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeHorizontal=5          # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeIn  =7          # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeInHorizontal=3          # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeInVertical=1          # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeLeft=23         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeNone=0          # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeOut =8          # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeOutHorizontal=4          # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeOutVertical=2          # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeRight=24         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeSpokes1=18         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeSpokes2=19         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeSpokes3=20         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeSpokes4=21         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeSpokes8=22         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeUp  =26         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeUpLeft=15         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeUpRight=17         # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectSubtypeVertical=6          # from enum MsoAnimFilterEffectSubtype
	msoAnimFilterEffectTypeBarn   =1          # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeBlinds =2          # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeBox    =3          # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeCheckerboard=4          # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeCircle =5          # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeDiamond=6          # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeDissolve=7          # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeFade   =8          # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeImage  =9          # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeNone   =0          # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypePixelate=10         # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypePlus   =11         # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeRandomBar=12         # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeSlide  =13         # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeStretch=14         # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeStrips =15         # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeWedge  =16         # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeWheel  =17         # from enum MsoAnimFilterEffectType
	msoAnimFilterEffectTypeWipe   =18         # from enum MsoAnimFilterEffectType
	msoAnimColor                  =7          # from enum MsoAnimProperty
	msoAnimHeight                 =4          # from enum MsoAnimProperty
	msoAnimNone                   =0          # from enum MsoAnimProperty
	msoAnimOpacity                =5          # from enum MsoAnimProperty
	msoAnimRotation               =6          # from enum MsoAnimProperty
	msoAnimShapeFillBackColor     =1007       # from enum MsoAnimProperty
	msoAnimShapeFillColor         =1005       # from enum MsoAnimProperty
	msoAnimShapeFillOn            =1004       # from enum MsoAnimProperty
	msoAnimShapeFillOpacity       =1006       # from enum MsoAnimProperty
	msoAnimShapeLineColor         =1009       # from enum MsoAnimProperty
	msoAnimShapeLineOn            =1008       # from enum MsoAnimProperty
	msoAnimShapePictureBrightness =1001       # from enum MsoAnimProperty
	msoAnimShapePictureContrast   =1000       # from enum MsoAnimProperty
	msoAnimShapePictureGamma      =1002       # from enum MsoAnimProperty
	msoAnimShapePictureGrayscale  =1003       # from enum MsoAnimProperty
	msoAnimShapeShadowColor       =1012       # from enum MsoAnimProperty
	msoAnimShapeShadowOffsetX     =1014       # from enum MsoAnimProperty
	msoAnimShapeShadowOffsetY     =1015       # from enum MsoAnimProperty
	msoAnimShapeShadowOn          =1010       # from enum MsoAnimProperty
	msoAnimShapeShadowOpacity     =1013       # from enum MsoAnimProperty
	msoAnimShapeShadowType        =1011       # from enum MsoAnimProperty
	msoAnimTextBulletCharacter    =111        # from enum MsoAnimProperty
	msoAnimTextBulletColor        =114        # from enum MsoAnimProperty
	msoAnimTextBulletFontName     =112        # from enum MsoAnimProperty
	msoAnimTextBulletNumber       =113        # from enum MsoAnimProperty
	msoAnimTextBulletRelativeSize =115        # from enum MsoAnimProperty
	msoAnimTextBulletStyle        =116        # from enum MsoAnimProperty
	msoAnimTextBulletType         =117        # from enum MsoAnimProperty
	msoAnimTextFontBold           =100        # from enum MsoAnimProperty
	msoAnimTextFontColor          =101        # from enum MsoAnimProperty
	msoAnimTextFontEmboss         =102        # from enum MsoAnimProperty
	msoAnimTextFontItalic         =103        # from enum MsoAnimProperty
	msoAnimTextFontName           =104        # from enum MsoAnimProperty
	msoAnimTextFontShadow         =105        # from enum MsoAnimProperty
	msoAnimTextFontSize           =106        # from enum MsoAnimProperty
	msoAnimTextFontStrikeThrough  =110        # from enum MsoAnimProperty
	msoAnimTextFontSubscript      =107        # from enum MsoAnimProperty
	msoAnimTextFontSuperscript    =108        # from enum MsoAnimProperty
	msoAnimTextFontUnderline      =109        # from enum MsoAnimProperty
	msoAnimVisibility             =8          # from enum MsoAnimProperty
	msoAnimWidth                  =3          # from enum MsoAnimProperty
	msoAnimX                      =1          # from enum MsoAnimProperty
	msoAnimY                      =2          # from enum MsoAnimProperty
	msoAnimTextUnitEffectByCharacter=1          # from enum MsoAnimTextUnitEffect
	msoAnimTextUnitEffectByParagraph=0          # from enum MsoAnimTextUnitEffect
	msoAnimTextUnitEffectByWord   =2          # from enum MsoAnimTextUnitEffect
	msoAnimTextUnitEffectMixed    =-1         # from enum MsoAnimTextUnitEffect
	msoAnimTriggerAfterPrevious   =3          # from enum MsoAnimTriggerType
	msoAnimTriggerMixed           =-1         # from enum MsoAnimTriggerType
	msoAnimTriggerNone            =0          # from enum MsoAnimTriggerType
	msoAnimTriggerOnPageClick     =1          # from enum MsoAnimTriggerType
	msoAnimTriggerOnShapeClick    =4          # from enum MsoAnimTriggerType
	msoAnimTriggerWithPrevious    =2          # from enum MsoAnimTriggerType
	msoAnimTypeColor              =2          # from enum MsoAnimType
	msoAnimTypeCommand            =6          # from enum MsoAnimType
	msoAnimTypeFilter             =7          # from enum MsoAnimType
	msoAnimTypeMixed              =-2         # from enum MsoAnimType
	msoAnimTypeMotion             =1          # from enum MsoAnimType
	msoAnimTypeNone               =0          # from enum MsoAnimType
	msoAnimTypeProperty           =5          # from enum MsoAnimType
	msoAnimTypeRotation           =4          # from enum MsoAnimType
	msoAnimTypeScale              =3          # from enum MsoAnimType
	msoAnimTypeSet                =8          # from enum MsoAnimType
	msoAnimateChartAllAtOnce      =7          # from enum MsoAnimateByLevel
	msoAnimateChartByCategory     =8          # from enum MsoAnimateByLevel
	msoAnimateChartByCategoryElements=9          # from enum MsoAnimateByLevel
	msoAnimateChartBySeries       =10         # from enum MsoAnimateByLevel
	msoAnimateChartBySeriesElements=11         # from enum MsoAnimateByLevel
	msoAnimateDiagramAllAtOnce    =12         # from enum MsoAnimateByLevel
	msoAnimateDiagramBreadthByLevel=16         # from enum MsoAnimateByLevel
	msoAnimateDiagramBreadthByNode=15         # from enum MsoAnimateByLevel
	msoAnimateDiagramClockwise    =17         # from enum MsoAnimateByLevel
	msoAnimateDiagramClockwiseIn  =18         # from enum MsoAnimateByLevel
	msoAnimateDiagramClockwiseOut =19         # from enum MsoAnimateByLevel
	msoAnimateDiagramCounterClockwise=20         # from enum MsoAnimateByLevel
	msoAnimateDiagramCounterClockwiseIn=21         # from enum MsoAnimateByLevel
	msoAnimateDiagramCounterClockwiseOut=22         # from enum MsoAnimateByLevel
	msoAnimateDiagramDepthByBranch=14         # from enum MsoAnimateByLevel
	msoAnimateDiagramDepthByNode  =13         # from enum MsoAnimateByLevel
	msoAnimateDiagramDown         =26         # from enum MsoAnimateByLevel
	msoAnimateDiagramInByRing     =23         # from enum MsoAnimateByLevel
	msoAnimateDiagramOutByRing    =24         # from enum MsoAnimateByLevel
	msoAnimateDiagramUp           =25         # from enum MsoAnimateByLevel
	msoAnimateLevelMixed          =-1         # from enum MsoAnimateByLevel
	msoAnimateLevelNone           =0          # from enum MsoAnimateByLevel
	msoAnimateTextByAllLevels     =1          # from enum MsoAnimateByLevel
	msoAnimateTextByFifthLevel    =6          # from enum MsoAnimateByLevel
	msoAnimateTextByFirstLevel    =2          # from enum MsoAnimateByLevel
	msoAnimateTextByFourthLevel   =5          # from enum MsoAnimateByLevel
	msoAnimateTextBySecondLevel   =3          # from enum MsoAnimateByLevel
	msoAnimateTextByThirdLevel    =4          # from enum MsoAnimateByLevel
	msoClickStateAfterAllAnimations=-2         # from enum MsoClickState
	msoClickStateBeforeAutomaticAnimations=-1         # from enum MsoClickState
	ppActionEndShow               =6          # from enum PpActionType
	ppActionFirstSlide            =3          # from enum PpActionType
	ppActionHyperlink             =7          # from enum PpActionType
	ppActionLastSlide             =4          # from enum PpActionType
	ppActionLastSlideViewed       =5          # from enum PpActionType
	ppActionMixed                 =-2         # from enum PpActionType
	ppActionNamedSlideShow        =10         # from enum PpActionType
	ppActionNextSlide             =1          # from enum PpActionType
	ppActionNone                  =0          # from enum PpActionType
	ppActionOLEVerb               =11         # from enum PpActionType
	ppActionPlay                  =12         # from enum PpActionType
	ppActionPreviousSlide         =2          # from enum PpActionType
	ppActionRunMacro              =8          # from enum PpActionType
	ppActionRunProgram            =9          # from enum PpActionType
	ppAdvanceModeMixed            =-2         # from enum PpAdvanceMode
	ppAdvanceOnClick              =1          # from enum PpAdvanceMode
	ppAdvanceOnTime               =2          # from enum PpAdvanceMode
	ppAfterEffectDim              =2          # from enum PpAfterEffect
	ppAfterEffectHide             =1          # from enum PpAfterEffect
	ppAfterEffectHideOnClick      =3          # from enum PpAfterEffect
	ppAfterEffectMixed            =-2         # from enum PpAfterEffect
	ppAfterEffectNothing          =0          # from enum PpAfterEffect
	ppAlertsAll                   =2          # from enum PpAlertLevel
	ppAlertsNone                  =1          # from enum PpAlertLevel
	ppArrangeCascade              =2          # from enum PpArrangeStyle
	ppArrangeTiled                =1          # from enum PpArrangeStyle
	ppAutoSizeMixed               =-2         # from enum PpAutoSize
	ppAutoSizeNone                =0          # from enum PpAutoSize
	ppAutoSizeShapeToFitText      =1          # from enum PpAutoSize
	ppBaselineAlignAuto           =5          # from enum PpBaselineAlignment
	ppBaselineAlignBaseline       =1          # from enum PpBaselineAlignment
	ppBaselineAlignCenter         =3          # from enum PpBaselineAlignment
	ppBaselineAlignFarEast50      =4          # from enum PpBaselineAlignment
	ppBaselineAlignMixed          =-2         # from enum PpBaselineAlignment
	ppBaselineAlignTop            =2          # from enum PpBaselineAlignment
	ppBorderBottom                =3          # from enum PpBorderType
	ppBorderDiagonalDown          =5          # from enum PpBorderType
	ppBorderDiagonalUp            =6          # from enum PpBorderType
	ppBorderLeft                  =2          # from enum PpBorderType
	ppBorderRight                 =4          # from enum PpBorderType
	ppBorderTop                   =1          # from enum PpBorderType
	ppBulletMixed                 =-2         # from enum PpBulletType
	ppBulletNone                  =0          # from enum PpBulletType
	ppBulletNumbered              =2          # from enum PpBulletType
	ppBulletPicture               =3          # from enum PpBulletType
	ppBulletUnnumbered            =1          # from enum PpBulletType
	ppCaseLower                   =2          # from enum PpChangeCase
	ppCaseSentence                =1          # from enum PpChangeCase
	ppCaseTitle                   =4          # from enum PpChangeCase
	ppCaseToggle                  =5          # from enum PpChangeCase
	ppCaseUpper                   =3          # from enum PpChangeCase
	ppAnimateByCategory           =2          # from enum PpChartUnitEffect
	ppAnimateByCategoryElements   =4          # from enum PpChartUnitEffect
	ppAnimateBySeries             =1          # from enum PpChartUnitEffect
	ppAnimateBySeriesElements     =3          # from enum PpChartUnitEffect
	ppAnimateChartAllAtOnce       =5          # from enum PpChartUnitEffect
	ppAnimateChartMixed           =-2         # from enum PpChartUnitEffect
	ppCheckInMajorVersion         =1          # from enum PpCheckInVersionType
	ppCheckInMinorVersion         =0          # from enum PpCheckInVersionType
	ppCheckInOverwriteVersion     =2          # from enum PpCheckInVersionType
	ppAccent1                     =6          # from enum PpColorSchemeIndex
	ppAccent2                     =7          # from enum PpColorSchemeIndex
	ppAccent3                     =8          # from enum PpColorSchemeIndex
	ppBackground                  =1          # from enum PpColorSchemeIndex
	ppFill                        =5          # from enum PpColorSchemeIndex
	ppForeground                  =2          # from enum PpColorSchemeIndex
	ppNotSchemeColor              =0          # from enum PpColorSchemeIndex
	ppSchemeColorMixed            =-2         # from enum PpColorSchemeIndex
	ppShadow                      =3          # from enum PpColorSchemeIndex
	ppTitle                       =4          # from enum PpColorSchemeIndex
	ppDateTimeFigureOut           =14         # from enum PpDateTimeFormat
	ppDateTimeFormatMixed         =-2         # from enum PpDateTimeFormat
	ppDateTimeHmm                 =10         # from enum PpDateTimeFormat
	ppDateTimeHmmss               =11         # from enum PpDateTimeFormat
	ppDateTimeMMMMdyyyy           =4          # from enum PpDateTimeFormat
	ppDateTimeMMMMyy              =6          # from enum PpDateTimeFormat
	ppDateTimeMMddyyHmm           =8          # from enum PpDateTimeFormat
	ppDateTimeMMddyyhmmAMPM       =9          # from enum PpDateTimeFormat
	ppDateTimeMMyy                =7          # from enum PpDateTimeFormat
	ppDateTimeMdyy                =1          # from enum PpDateTimeFormat
	ppDateTimedMMMMyyyy           =3          # from enum PpDateTimeFormat
	ppDateTimedMMMyy              =5          # from enum PpDateTimeFormat
	ppDateTimeddddMMMMddyyyy      =2          # from enum PpDateTimeFormat
	ppDateTimehmmAMPM             =12         # from enum PpDateTimeFormat
	ppDateTimehmmssAMPM           =13         # from enum PpDateTimeFormat
	ppDirectionLeftToRight        =1          # from enum PpDirection
	ppDirectionMixed              =-2         # from enum PpDirection
	ppDirectionRightToLeft        =2          # from enum PpDirection
	ppEffectAppear                =3844       # from enum PpEntryEffect
	ppEffectBlindsHorizontal      =769        # from enum PpEntryEffect
	ppEffectBlindsVertical        =770        # from enum PpEntryEffect
	ppEffectBoxIn                 =3074       # from enum PpEntryEffect
	ppEffectBoxOut                =3073       # from enum PpEntryEffect
	ppEffectCheckerboardAcross    =1025       # from enum PpEntryEffect
	ppEffectCheckerboardDown      =1026       # from enum PpEntryEffect
	ppEffectCircleOut             =3845       # from enum PpEntryEffect
	ppEffectCombHorizontal        =3847       # from enum PpEntryEffect
	ppEffectCombVertical          =3848       # from enum PpEntryEffect
	ppEffectCoverDown             =1284       # from enum PpEntryEffect
	ppEffectCoverLeft             =1281       # from enum PpEntryEffect
	ppEffectCoverLeftDown         =1287       # from enum PpEntryEffect
	ppEffectCoverLeftUp           =1285       # from enum PpEntryEffect
	ppEffectCoverRight            =1283       # from enum PpEntryEffect
	ppEffectCoverRightDown        =1288       # from enum PpEntryEffect
	ppEffectCoverRightUp          =1286       # from enum PpEntryEffect
	ppEffectCoverUp               =1282       # from enum PpEntryEffect
	ppEffectCrawlFromDown         =3344       # from enum PpEntryEffect
	ppEffectCrawlFromLeft         =3341       # from enum PpEntryEffect
	ppEffectCrawlFromRight        =3343       # from enum PpEntryEffect
	ppEffectCrawlFromUp           =3342       # from enum PpEntryEffect
	ppEffectCut                   =257        # from enum PpEntryEffect
	ppEffectCutThroughBlack       =258        # from enum PpEntryEffect
	ppEffectDiamondOut            =3846       # from enum PpEntryEffect
	ppEffectDissolve              =1537       # from enum PpEntryEffect
	ppEffectFade                  =1793       # from enum PpEntryEffect
	ppEffectFadeSmoothly          =3849       # from enum PpEntryEffect
	ppEffectFlashOnceFast         =3841       # from enum PpEntryEffect
	ppEffectFlashOnceMedium       =3842       # from enum PpEntryEffect
	ppEffectFlashOnceSlow         =3843       # from enum PpEntryEffect
	ppEffectFlyFromBottom         =3332       # from enum PpEntryEffect
	ppEffectFlyFromBottomLeft     =3335       # from enum PpEntryEffect
	ppEffectFlyFromBottomRight    =3336       # from enum PpEntryEffect
	ppEffectFlyFromLeft           =3329       # from enum PpEntryEffect
	ppEffectFlyFromRight          =3331       # from enum PpEntryEffect
	ppEffectFlyFromTop            =3330       # from enum PpEntryEffect
	ppEffectFlyFromTopLeft        =3333       # from enum PpEntryEffect
	ppEffectFlyFromTopRight       =3334       # from enum PpEntryEffect
	ppEffectMixed                 =-2         # from enum PpEntryEffect
	ppEffectNewsflash             =3850       # from enum PpEntryEffect
	ppEffectNone                  =0          # from enum PpEntryEffect
	ppEffectPeekFromDown          =3338       # from enum PpEntryEffect
	ppEffectPeekFromLeft          =3337       # from enum PpEntryEffect
	ppEffectPeekFromRight         =3339       # from enum PpEntryEffect
	ppEffectPeekFromUp            =3340       # from enum PpEntryEffect
	ppEffectPlusOut               =3851       # from enum PpEntryEffect
	ppEffectPushDown              =3852       # from enum PpEntryEffect
	ppEffectPushLeft              =3853       # from enum PpEntryEffect
	ppEffectPushRight             =3854       # from enum PpEntryEffect
	ppEffectPushUp                =3855       # from enum PpEntryEffect
	ppEffectRandom                =513        # from enum PpEntryEffect
	ppEffectRandomBarsHorizontal  =2305       # from enum PpEntryEffect
	ppEffectRandomBarsVertical    =2306       # from enum PpEntryEffect
	ppEffectSpiral                =3357       # from enum PpEntryEffect
	ppEffectSplitHorizontalIn     =3586       # from enum PpEntryEffect
	ppEffectSplitHorizontalOut    =3585       # from enum PpEntryEffect
	ppEffectSplitVerticalIn       =3588       # from enum PpEntryEffect
	ppEffectSplitVerticalOut      =3587       # from enum PpEntryEffect
	ppEffectStretchAcross         =3351       # from enum PpEntryEffect
	ppEffectStretchDown           =3355       # from enum PpEntryEffect
	ppEffectStretchLeft           =3352       # from enum PpEntryEffect
	ppEffectStretchRight          =3354       # from enum PpEntryEffect
	ppEffectStretchUp             =3353       # from enum PpEntryEffect
	ppEffectStripsDownLeft        =2563       # from enum PpEntryEffect
	ppEffectStripsDownRight       =2564       # from enum PpEntryEffect
	ppEffectStripsLeftDown        =2567       # from enum PpEntryEffect
	ppEffectStripsLeftUp          =2565       # from enum PpEntryEffect
	ppEffectStripsRightDown       =2568       # from enum PpEntryEffect
	ppEffectStripsRightUp         =2566       # from enum PpEntryEffect
	ppEffectStripsUpLeft          =2561       # from enum PpEntryEffect
	ppEffectStripsUpRight         =2562       # from enum PpEntryEffect
	ppEffectSwivel                =3356       # from enum PpEntryEffect
	ppEffectUncoverDown           =2052       # from enum PpEntryEffect
	ppEffectUncoverLeft           =2049       # from enum PpEntryEffect
	ppEffectUncoverLeftDown       =2055       # from enum PpEntryEffect
	ppEffectUncoverLeftUp         =2053       # from enum PpEntryEffect
	ppEffectUncoverRight          =2051       # from enum PpEntryEffect
	ppEffectUncoverRightDown      =2056       # from enum PpEntryEffect
	ppEffectUncoverRightUp        =2054       # from enum PpEntryEffect
	ppEffectUncoverUp             =2050       # from enum PpEntryEffect
	ppEffectWedge                 =3856       # from enum PpEntryEffect
	ppEffectWheel1Spoke           =3857       # from enum PpEntryEffect
	ppEffectWheel2Spokes          =3858       # from enum PpEntryEffect
	ppEffectWheel3Spokes          =3859       # from enum PpEntryEffect
	ppEffectWheel4Spokes          =3860       # from enum PpEntryEffect
	ppEffectWheel8Spokes          =3861       # from enum PpEntryEffect
	ppEffectWipeDown              =2820       # from enum PpEntryEffect
	ppEffectWipeLeft              =2817       # from enum PpEntryEffect
	ppEffectWipeRight             =2819       # from enum PpEntryEffect
	ppEffectWipeUp                =2818       # from enum PpEntryEffect
	ppEffectZoomBottom            =3350       # from enum PpEntryEffect
	ppEffectZoomCenter            =3349       # from enum PpEntryEffect
	ppEffectZoomIn                =3345       # from enum PpEntryEffect
	ppEffectZoomInSlightly        =3346       # from enum PpEntryEffect
	ppEffectZoomOut               =3347       # from enum PpEntryEffect
	ppEffectZoomOutSlightly       =3348       # from enum PpEntryEffect
	ppClipRelativeToSlide         =2          # from enum PpExportMode
	ppRelativeToSlide             =1          # from enum PpExportMode
	ppScaleToFit                  =3          # from enum PpExportMode
	ppScaleXY                     =4          # from enum PpExportMode
	ppFarEastLineBreakLevelCustom =3          # from enum PpFarEastLineBreakLevel
	ppFarEastLineBreakLevelNormal =1          # from enum PpFarEastLineBreakLevel
	ppFarEastLineBreakLevelStrict =2          # from enum PpFarEastLineBreakLevel
	ppFileDialogOpen              =1          # from enum PpFileDialogType
	ppFileDialogSave              =2          # from enum PpFileDialogType
	ppFixedFormatIntentPrint      =2          # from enum PpFixedFormatIntent
	ppFixedFormatIntentScreen     =1          # from enum PpFixedFormatIntent
	ppFixedFormatTypePDF          =2          # from enum PpFixedFormatType
	ppFixedFormatTypeXPS          =1          # from enum PpFixedFormatType
	ppFollowColorsMixed           =-2         # from enum PpFollowColors
	ppFollowColorsNone            =0          # from enum PpFollowColors
	ppFollowColorsScheme          =1          # from enum PpFollowColors
	ppFollowColorsTextAndBackground=2          # from enum PpFollowColors
	ppFrameColorsBlackTextOnWhite =5          # from enum PpFrameColors
	ppFrameColorsBrowserColors    =1          # from enum PpFrameColors
	ppFrameColorsPresentationSchemeAccentColor=3          # from enum PpFrameColors
	ppFrameColorsPresentationSchemeTextColor=2          # from enum PpFrameColors
	ppFrameColorsWhiteTextOnBlack =4          # from enum PpFrameColors
	ppHTMLAutodetect              =4          # from enum PpHTMLVersion
	ppHTMLDual                    =3          # from enum PpHTMLVersion
	ppHTMLv3                      =1          # from enum PpHTMLVersion
	ppHTMLv4                      =2          # from enum PpHTMLVersion
	ppIndentControlMixed          =-2         # from enum PpIndentControl
	ppIndentKeepAttr              =2          # from enum PpIndentControl
	ppIndentReplaceAttr           =1          # from enum PpIndentControl
	ppMediaTypeMixed              =-2         # from enum PpMediaType
	ppMediaTypeMovie              =3          # from enum PpMediaType
	ppMediaTypeOther              =1          # from enum PpMediaType
	ppMediaTypeSound              =2          # from enum PpMediaType
	ppMouseClick                  =1          # from enum PpMouseActivation
	ppMouseOver                   =2          # from enum PpMouseActivation
	ppBulletAlphaLCParenBoth      =8          # from enum PpNumberedBulletStyle
	ppBulletAlphaLCParenRight     =9          # from enum PpNumberedBulletStyle
	ppBulletAlphaLCPeriod         =0          # from enum PpNumberedBulletStyle
	ppBulletAlphaUCParenBoth      =10         # from enum PpNumberedBulletStyle
	ppBulletAlphaUCParenRight     =11         # from enum PpNumberedBulletStyle
	ppBulletAlphaUCPeriod         =1          # from enum PpNumberedBulletStyle
	ppBulletArabicAbjadDash       =24         # from enum PpNumberedBulletStyle
	ppBulletArabicAlphaDash       =23         # from enum PpNumberedBulletStyle
	ppBulletArabicDBPeriod        =29         # from enum PpNumberedBulletStyle
	ppBulletArabicDBPlain         =28         # from enum PpNumberedBulletStyle
	ppBulletArabicParenBoth       =12         # from enum PpNumberedBulletStyle
	ppBulletArabicParenRight      =2          # from enum PpNumberedBulletStyle
	ppBulletArabicPeriod          =3          # from enum PpNumberedBulletStyle
	ppBulletArabicPlain           =13         # from enum PpNumberedBulletStyle
	ppBulletCircleNumDBPlain      =18         # from enum PpNumberedBulletStyle
	ppBulletCircleNumWDBlackPlain =20         # from enum PpNumberedBulletStyle
	ppBulletCircleNumWDWhitePlain =19         # from enum PpNumberedBulletStyle
	ppBulletHebrewAlphaDash       =25         # from enum PpNumberedBulletStyle
	ppBulletHindiAlpha1Period     =40         # from enum PpNumberedBulletStyle
	ppBulletHindiAlphaPeriod      =36         # from enum PpNumberedBulletStyle
	ppBulletHindiNumParenRight    =39         # from enum PpNumberedBulletStyle
	ppBulletHindiNumPeriod        =37         # from enum PpNumberedBulletStyle
	ppBulletKanjiKoreanPeriod     =27         # from enum PpNumberedBulletStyle
	ppBulletKanjiKoreanPlain      =26         # from enum PpNumberedBulletStyle
	ppBulletKanjiSimpChinDBPeriod =38         # from enum PpNumberedBulletStyle
	ppBulletRomanLCParenBoth      =4          # from enum PpNumberedBulletStyle
	ppBulletRomanLCParenRight     =5          # from enum PpNumberedBulletStyle
	ppBulletRomanLCPeriod         =6          # from enum PpNumberedBulletStyle
	ppBulletRomanUCParenBoth      =14         # from enum PpNumberedBulletStyle
	ppBulletRomanUCParenRight     =15         # from enum PpNumberedBulletStyle
	ppBulletRomanUCPeriod         =7          # from enum PpNumberedBulletStyle
	ppBulletSimpChinPeriod        =17         # from enum PpNumberedBulletStyle
	ppBulletSimpChinPlain         =16         # from enum PpNumberedBulletStyle
	ppBulletStyleMixed            =-2         # from enum PpNumberedBulletStyle
	ppBulletThaiAlphaParenBoth    =32         # from enum PpNumberedBulletStyle
	ppBulletThaiAlphaParenRight   =31         # from enum PpNumberedBulletStyle
	ppBulletThaiAlphaPeriod       =30         # from enum PpNumberedBulletStyle
	ppBulletThaiNumParenBoth      =35         # from enum PpNumberedBulletStyle
	ppBulletThaiNumParenRight     =34         # from enum PpNumberedBulletStyle
	ppBulletThaiNumPeriod         =33         # from enum PpNumberedBulletStyle
	ppBulletTradChinPeriod        =22         # from enum PpNumberedBulletStyle
	ppBulletTradChinPlain         =21         # from enum PpNumberedBulletStyle
	ppAlignCenter                 =2          # from enum PpParagraphAlignment
	ppAlignDistribute             =5          # from enum PpParagraphAlignment
	ppAlignJustify                =4          # from enum PpParagraphAlignment
	ppAlignJustifyLow             =7          # from enum PpParagraphAlignment
	ppAlignLeft                   =1          # from enum PpParagraphAlignment
	ppAlignRight                  =3          # from enum PpParagraphAlignment
	ppAlignThaiDistribute         =6          # from enum PpParagraphAlignment
	ppAlignmentMixed              =-2         # from enum PpParagraphAlignment
	ppPasteBitmap                 =1          # from enum PpPasteDataType
	ppPasteDefault                =0          # from enum PpPasteDataType
	ppPasteEnhancedMetafile       =2          # from enum PpPasteDataType
	ppPasteGIF                    =4          # from enum PpPasteDataType
	ppPasteHTML                   =8          # from enum PpPasteDataType
	ppPasteJPG                    =5          # from enum PpPasteDataType
	ppPasteMetafilePicture        =3          # from enum PpPasteDataType
	ppPasteOLEObject              =10         # from enum PpPasteDataType
	ppPastePNG                    =6          # from enum PpPasteDataType
	ppPasteRTF                    =9          # from enum PpPasteDataType
	ppPasteShape                  =11         # from enum PpPasteDataType
	ppPasteText                   =7          # from enum PpPasteDataType
	ppPlaceholderBitmap           =9          # from enum PpPlaceholderType
	ppPlaceholderBody             =2          # from enum PpPlaceholderType
	ppPlaceholderCenterTitle      =3          # from enum PpPlaceholderType
	ppPlaceholderChart            =8          # from enum PpPlaceholderType
	ppPlaceholderDate             =16         # from enum PpPlaceholderType
	ppPlaceholderFooter           =15         # from enum PpPlaceholderType
	ppPlaceholderHeader           =14         # from enum PpPlaceholderType
	ppPlaceholderMediaClip        =10         # from enum PpPlaceholderType
	ppPlaceholderMixed            =-2         # from enum PpPlaceholderType
	ppPlaceholderObject           =7          # from enum PpPlaceholderType
	ppPlaceholderOrgChart         =11         # from enum PpPlaceholderType
	ppPlaceholderPicture          =18         # from enum PpPlaceholderType
	ppPlaceholderSlideNumber      =13         # from enum PpPlaceholderType
	ppPlaceholderSubtitle         =4          # from enum PpPlaceholderType
	ppPlaceholderTable            =12         # from enum PpPlaceholderType
	ppPlaceholderTitle            =1          # from enum PpPlaceholderType
	ppPlaceholderVerticalBody     =6          # from enum PpPlaceholderType
	ppPlaceholderVerticalObject   =17         # from enum PpPlaceholderType
	ppPlaceholderVerticalTitle    =5          # from enum PpPlaceholderType
	ppPrintBlackAndWhite          =2          # from enum PpPrintColorType
	ppPrintColor                  =1          # from enum PpPrintColorType
	ppPrintPureBlackAndWhite      =3          # from enum PpPrintColorType
	ppPrintHandoutHorizontalFirst =2          # from enum PpPrintHandoutOrder
	ppPrintHandoutVerticalFirst   =1          # from enum PpPrintHandoutOrder
	ppPrintOutputBuildSlides      =7          # from enum PpPrintOutputType
	ppPrintOutputFourSlideHandouts=8          # from enum PpPrintOutputType
	ppPrintOutputNineSlideHandouts=9          # from enum PpPrintOutputType
	ppPrintOutputNotesPages       =5          # from enum PpPrintOutputType
	ppPrintOutputOneSlideHandouts =10         # from enum PpPrintOutputType
	ppPrintOutputOutline          =6          # from enum PpPrintOutputType
	ppPrintOutputSixSlideHandouts =4          # from enum PpPrintOutputType
	ppPrintOutputSlides           =1          # from enum PpPrintOutputType
	ppPrintOutputThreeSlideHandouts=3          # from enum PpPrintOutputType
	ppPrintOutputTwoSlideHandouts =2          # from enum PpPrintOutputType
	ppPrintAll                    =1          # from enum PpPrintRangeType
	ppPrintCurrent                =3          # from enum PpPrintRangeType
	ppPrintNamedSlideShow         =5          # from enum PpPrintRangeType
	ppPrintSelection              =2          # from enum PpPrintRangeType
	ppPrintSlideRange             =4          # from enum PpPrintRangeType
	ppPublishAll                  =1          # from enum PpPublishSourceType
	ppPublishNamedSlideShow       =3          # from enum PpPublishSourceType
	ppPublishSlideRange           =2          # from enum PpPublishSourceType
	ppRDIAll                      =99         # from enum PpRemoveDocInfoType
	ppRDIComments                 =1          # from enum PpRemoveDocInfoType
	ppRDIContentType              =16         # from enum PpRemoveDocInfoType
	ppRDIDocumentManagementPolicy =15         # from enum PpRemoveDocInfoType
	ppRDIDocumentProperties       =8          # from enum PpRemoveDocInfoType
	ppRDIDocumentServerProperties =14         # from enum PpRemoveDocInfoType
	ppRDIDocumentWorkspace        =10         # from enum PpRemoveDocInfoType
	ppRDIInkAnnotations           =11         # from enum PpRemoveDocInfoType
	ppRDIPublishPath              =13         # from enum PpRemoveDocInfoType
	ppRDIRemovePersonalInformation=4          # from enum PpRemoveDocInfoType
	ppRDISlideUpdateInformation   =17         # from enum PpRemoveDocInfoType
	ppRevisionInfoBaseline        =1          # from enum PpRevisionInfo
	ppRevisionInfoMerged          =2          # from enum PpRevisionInfo
	ppRevisionInfoNone            =0          # from enum PpRevisionInfo
	ppSaveAsAddIn                 =8          # from enum PpSaveAsFileType
	ppSaveAsBMP                   =19         # from enum PpSaveAsFileType
	ppSaveAsDefault               =11         # from enum PpSaveAsFileType
	ppSaveAsEMF                   =23         # from enum PpSaveAsFileType
	ppSaveAsGIF                   =16         # from enum PpSaveAsFileType
	ppSaveAsHTML                  =12         # from enum PpSaveAsFileType
	ppSaveAsHTMLDual              =14         # from enum PpSaveAsFileType
	ppSaveAsHTMLv3                =13         # from enum PpSaveAsFileType
	ppSaveAsJPG                   =17         # from enum PpSaveAsFileType
	ppSaveAsMetaFile              =15         # from enum PpSaveAsFileType
	ppSaveAsOpenXMLAddin          =30         # from enum PpSaveAsFileType
	ppSaveAsOpenXMLPresentation   =24         # from enum PpSaveAsFileType
	ppSaveAsOpenXMLPresentationMacroEnabled=25         # from enum PpSaveAsFileType
	ppSaveAsOpenXMLShow           =28         # from enum PpSaveAsFileType
	ppSaveAsOpenXMLShowMacroEnabled=29         # from enum PpSaveAsFileType
	ppSaveAsOpenXMLTemplate       =26         # from enum PpSaveAsFileType
	ppSaveAsOpenXMLTemplateMacroEnabled=27         # from enum PpSaveAsFileType
	ppSaveAsOpenXMLTheme          =31         # from enum PpSaveAsFileType
	ppSaveAsPDF                   =32         # from enum PpSaveAsFileType
	ppSaveAsPNG                   =18         # from enum PpSaveAsFileType
	ppSaveAsPowerPoint3           =4          # from enum PpSaveAsFileType
	ppSaveAsPowerPoint4           =3          # from enum PpSaveAsFileType
	ppSaveAsPowerPoint4FarEast    =10         # from enum PpSaveAsFileType
	ppSaveAsPowerPoint7           =2          # from enum PpSaveAsFileType
	ppSaveAsPresForReview         =22         # from enum PpSaveAsFileType
	ppSaveAsPresentation          =1          # from enum PpSaveAsFileType
	ppSaveAsRTF                   =6          # from enum PpSaveAsFileType
	ppSaveAsShow                  =7          # from enum PpSaveAsFileType
	ppSaveAsTIF                   =21         # from enum PpSaveAsFileType
	ppSaveAsTemplate              =5          # from enum PpSaveAsFileType
	ppSaveAsWebArchive            =20         # from enum PpSaveAsFileType
	ppSaveAsXMLPresentation       =34         # from enum PpSaveAsFileType
	ppSaveAsXPS                   =33         # from enum PpSaveAsFileType
	ppSelectionNone               =0          # from enum PpSelectionType
	ppSelectionShapes             =2          # from enum PpSelectionType
	ppSelectionSlides             =1          # from enum PpSelectionType
	ppSelectionText               =3          # from enum PpSelectionType
	ppShapeFormatBMP              =3          # from enum PpShapeFormat
	ppShapeFormatEMF              =5          # from enum PpShapeFormat
	ppShapeFormatGIF              =0          # from enum PpShapeFormat
	ppShapeFormatJPG              =1          # from enum PpShapeFormat
	ppShapeFormatPNG              =2          # from enum PpShapeFormat
	ppShapeFormatWMF              =4          # from enum PpShapeFormat
	ppLayoutBlank                 =12         # from enum PpSlideLayout
	ppLayoutChart                 =8          # from enum PpSlideLayout
	ppLayoutChartAndText          =6          # from enum PpSlideLayout
	ppLayoutClipArtAndVerticalText=26         # from enum PpSlideLayout
	ppLayoutClipartAndText        =10         # from enum PpSlideLayout
	ppLayoutComparison            =34         # from enum PpSlideLayout
	ppLayoutContentWithCaption    =35         # from enum PpSlideLayout
	ppLayoutCustom                =32         # from enum PpSlideLayout
	ppLayoutFourObjects           =24         # from enum PpSlideLayout
	ppLayoutLargeObject           =15         # from enum PpSlideLayout
	ppLayoutMediaClipAndText      =18         # from enum PpSlideLayout
	ppLayoutMixed                 =-2         # from enum PpSlideLayout
	ppLayoutObject                =16         # from enum PpSlideLayout
	ppLayoutObjectAndText         =14         # from enum PpSlideLayout
	ppLayoutObjectAndTwoObjects   =30         # from enum PpSlideLayout
	ppLayoutObjectOverText        =19         # from enum PpSlideLayout
	ppLayoutOrgchart              =7          # from enum PpSlideLayout
	ppLayoutPictureWithCaption    =36         # from enum PpSlideLayout
	ppLayoutSectionHeader         =33         # from enum PpSlideLayout
	ppLayoutTable                 =4          # from enum PpSlideLayout
	ppLayoutText                  =2          # from enum PpSlideLayout
	ppLayoutTextAndChart          =5          # from enum PpSlideLayout
	ppLayoutTextAndClipart        =9          # from enum PpSlideLayout
	ppLayoutTextAndMediaClip      =17         # from enum PpSlideLayout
	ppLayoutTextAndObject         =13         # from enum PpSlideLayout
	ppLayoutTextAndTwoObjects     =21         # from enum PpSlideLayout
	ppLayoutTextOverObject        =20         # from enum PpSlideLayout
	ppLayoutTitle                 =1          # from enum PpSlideLayout
	ppLayoutTitleOnly             =11         # from enum PpSlideLayout
	ppLayoutTwoColumnText         =3          # from enum PpSlideLayout
	ppLayoutTwoObjects            =29         # from enum PpSlideLayout
	ppLayoutTwoObjectsAndObject   =31         # from enum PpSlideLayout
	ppLayoutTwoObjectsAndText     =22         # from enum PpSlideLayout
	ppLayoutTwoObjectsOverText    =23         # from enum PpSlideLayout
	ppLayoutVerticalText          =25         # from enum PpSlideLayout
	ppLayoutVerticalTitleAndText  =27         # from enum PpSlideLayout
	ppLayoutVerticalTitleAndTextOverChart=28         # from enum PpSlideLayout
	ppSlideShowManualAdvance      =1          # from enum PpSlideShowAdvanceMode
	ppSlideShowRehearseNewTimings =3          # from enum PpSlideShowAdvanceMode
	ppSlideShowUseSlideTimings    =2          # from enum PpSlideShowAdvanceMode
	ppSlideShowPointerAlwaysHidden=3          # from enum PpSlideShowPointerType
	ppSlideShowPointerArrow       =1          # from enum PpSlideShowPointerType
	ppSlideShowPointerAutoArrow   =4          # from enum PpSlideShowPointerType
	ppSlideShowPointerEraser      =5          # from enum PpSlideShowPointerType
	ppSlideShowPointerNone        =0          # from enum PpSlideShowPointerType
	ppSlideShowPointerPen         =2          # from enum PpSlideShowPointerType
	ppShowAll                     =1          # from enum PpSlideShowRangeType
	ppShowNamedSlideShow          =3          # from enum PpSlideShowRangeType
	ppShowSlideRange              =2          # from enum PpSlideShowRangeType
	ppSlideShowBlackScreen        =3          # from enum PpSlideShowState
	ppSlideShowDone               =5          # from enum PpSlideShowState
	ppSlideShowPaused             =2          # from enum PpSlideShowState
	ppSlideShowRunning            =1          # from enum PpSlideShowState
	ppSlideShowWhiteScreen        =4          # from enum PpSlideShowState
	ppShowTypeKiosk               =3          # from enum PpSlideShowType
	ppShowTypeSpeaker             =1          # from enum PpSlideShowType
	ppShowTypeWindow              =2          # from enum PpSlideShowType
	ppSlideSize35MM               =4          # from enum PpSlideSizeType
	ppSlideSizeA3Paper            =9          # from enum PpSlideSizeType
	ppSlideSizeA4Paper            =3          # from enum PpSlideSizeType
	ppSlideSizeB4ISOPaper         =10         # from enum PpSlideSizeType
	ppSlideSizeB4JISPaper         =12         # from enum PpSlideSizeType
	ppSlideSizeB5ISOPaper         =11         # from enum PpSlideSizeType
	ppSlideSizeB5JISPaper         =13         # from enum PpSlideSizeType
	ppSlideSizeBanner             =6          # from enum PpSlideSizeType
	ppSlideSizeCustom             =7          # from enum PpSlideSizeType
	ppSlideSizeHagakiCard         =14         # from enum PpSlideSizeType
	ppSlideSizeLedgerPaper        =8          # from enum PpSlideSizeType
	ppSlideSizeLetterPaper        =2          # from enum PpSlideSizeType
	ppSlideSizeOnScreen           =1          # from enum PpSlideSizeType
	ppSlideSizeOnScreen16x10      =16         # from enum PpSlideSizeType
	ppSlideSizeOnScreen16x9       =15         # from enum PpSlideSizeType
	ppSlideSizeOverhead           =5          # from enum PpSlideSizeType
	ppSoundEffectsMixed           =-2         # from enum PpSoundEffectType
	ppSoundFile                   =2          # from enum PpSoundEffectType
	ppSoundNone                   =0          # from enum PpSoundEffectType
	ppSoundStopPrevious           =1          # from enum PpSoundEffectType
	ppSoundFormatCDAudio          =3          # from enum PpSoundFormatType
	ppSoundFormatMIDI             =2          # from enum PpSoundFormatType
	ppSoundFormatMixed            =-2         # from enum PpSoundFormatType
	ppSoundFormatNone             =0          # from enum PpSoundFormatType
	ppSoundFormatWAV              =1          # from enum PpSoundFormatType
	ppTabStopCenter               =2          # from enum PpTabStopType
	ppTabStopDecimal              =4          # from enum PpTabStopType
	ppTabStopLeft                 =1          # from enum PpTabStopType
	ppTabStopMixed                =-2         # from enum PpTabStopType
	ppTabStopRight                =3          # from enum PpTabStopType
	ppAnimateByAllLevels          =16         # from enum PpTextLevelEffect
	ppAnimateByFifthLevel         =5          # from enum PpTextLevelEffect
	ppAnimateByFirstLevel         =1          # from enum PpTextLevelEffect
	ppAnimateByFourthLevel        =4          # from enum PpTextLevelEffect
	ppAnimateBySecondLevel        =2          # from enum PpTextLevelEffect
	ppAnimateByThirdLevel         =3          # from enum PpTextLevelEffect
	ppAnimateLevelMixed           =-2         # from enum PpTextLevelEffect
	ppAnimateLevelNone            =0          # from enum PpTextLevelEffect
	ppBodyStyle                   =3          # from enum PpTextStyleType
	ppDefaultStyle                =1          # from enum PpTextStyleType
	ppTitleStyle                  =2          # from enum PpTextStyleType
	ppAnimateByCharacter          =2          # from enum PpTextUnitEffect
	ppAnimateByParagraph          =0          # from enum PpTextUnitEffect
	ppAnimateByWord               =1          # from enum PpTextUnitEffect
	ppAnimateUnitMixed            =-2         # from enum PpTextUnitEffect
	ppTransitionSpeedFast         =3          # from enum PpTransitionSpeed
	ppTransitionSpeedMedium       =2          # from enum PpTransitionSpeed
	ppTransitionSpeedMixed        =-2         # from enum PpTransitionSpeed
	ppTransitionSpeedSlow         =1          # from enum PpTransitionSpeed
	ppUpdateOptionAutomatic       =2          # from enum PpUpdateOption
	ppUpdateOptionManual          =1          # from enum PpUpdateOption
	ppUpdateOptionMixed           =-2         # from enum PpUpdateOption
	ppViewHandoutMaster           =4          # from enum PpViewType
	ppViewMasterThumbnails        =12         # from enum PpViewType
	ppViewNormal                  =9          # from enum PpViewType
	ppViewNotesMaster             =5          # from enum PpViewType
	ppViewNotesPage               =3          # from enum PpViewType
	ppViewOutline                 =6          # from enum PpViewType
	ppViewPrintPreview            =10         # from enum PpViewType
	ppViewSlide                   =1          # from enum PpViewType
	ppViewSlideMaster             =2          # from enum PpViewType
	ppViewSlideSorter             =7          # from enum PpViewType
	ppViewThumbnails              =11         # from enum PpViewType
	ppViewTitleMaster             =8          # from enum PpViewType
	ppWindowMaximized             =3          # from enum PpWindowState
	ppWindowMinimized             =2          # from enum PpWindowState
	ppWindowNormal                =1          # from enum PpWindowState

RecordMap = {
}

CLSIDToClassMap = {}
CLSIDToPackageMap = {
	'{914934E8-5A91-11CF-8700-00AA0060263B}' : u'RotationEffect',
	'{914934E9-5A91-11CF-8700-00AA0060263B}' : u'PropertyEffect',
	'{914934EA-5A91-11CF-8700-00AA0060263B}' : u'AnimationPoints',
	'{914934EB-5A91-11CF-8700-00AA0060263B}' : u'AnimationPoint',
	'{914934EC-5A91-11CF-8700-00AA0060263B}' : u'CanvasShapes',
	'{914934ED-5A91-11CF-8700-00AA0060263B}' : u'AutoCorrect',
	'{914934EE-5A91-11CF-8700-00AA0060263B}' : u'Options',
	'{914934EF-5A91-11CF-8700-00AA0060263B}' : u'CommandEffect',
	'{914934F0-5A91-11CF-8700-00AA0060263B}' : u'FilterEffect',
	'{914934F1-5A91-11CF-8700-00AA0060263B}' : u'SetEffect',
	'{914934F2-5A91-11CF-8700-00AA0060263B}' : u'CustomLayouts',
	'{914934F3-5A91-11CF-8700-00AA0060263B}' : u'CustomLayout',
	'{914934F5-5A91-11CF-8700-00AA0060263B}' : u'TableStyle',
	'{914934F6-5A91-11CF-8700-00AA0060263B}' : u'CustomerData',
	'{914934F7-5A91-11CF-8700-00AA0060263B}' : u'Research',
	'{914934F8-5A91-11CF-8700-00AA0060263B}' : u'TableBackground',
	'{914934F9-5A91-11CF-8700-00AA0060263B}' : u'TextFrame2',
	'{91493441-5A91-11CF-8700-00AA0060263B}' : u'Application',
	'{91493442-5A91-11CF-8700-00AA0060263B}' : u'_Application',
	'{91493443-5A91-11CF-8700-00AA0060263B}' : u'Global',
	'{91493444-5A91-11CF-8700-00AA0060263B}' : u'Presentation',
	'{91493445-5A91-11CF-8700-00AA0060263B}' : u'Slide',
	'{91493446-5A91-11CF-8700-00AA0060263B}' : u'OLEControl',
	'{91493447-5A91-11CF-8700-00AA0060263B}' : u'Master',
	'{91493448-5A91-11CF-8700-00AA0060263B}' : u'PowerRex',
	'{91493450-5A91-11CF-8700-00AA0060263B}' : u'Collection',
	'{91493451-5A91-11CF-8700-00AA0060263B}' : u'_Global',
	'{91493452-5A91-11CF-8700-00AA0060263B}' : u'ColorFormat',
	'{91493453-5A91-11CF-8700-00AA0060263B}' : u'SlideShowWindow',
	'{91493454-5A91-11CF-8700-00AA0060263B}' : u'Selection',
	'{91493455-5A91-11CF-8700-00AA0060263B}' : u'DocumentWindows',
	'{91493456-5A91-11CF-8700-00AA0060263B}' : u'SlideShowWindows',
	'{91493457-5A91-11CF-8700-00AA0060263B}' : u'DocumentWindow',
	'{91493458-5A91-11CF-8700-00AA0060263B}' : u'View',
	'{91493459-5A91-11CF-8700-00AA0060263B}' : u'SlideShowView',
	'{9149345A-5A91-11CF-8700-00AA0060263B}' : u'SlideShowSettings',
	'{9149345B-5A91-11CF-8700-00AA0060263B}' : u'NamedSlideShows',
	'{9149345C-5A91-11CF-8700-00AA0060263B}' : u'NamedSlideShow',
	'{9149345D-5A91-11CF-8700-00AA0060263B}' : u'PrintOptions',
	'{9149345E-5A91-11CF-8700-00AA0060263B}' : u'PrintRanges',
	'{9149345F-5A91-11CF-8700-00AA0060263B}' : u'PrintRange',
	'{91493460-5A91-11CF-8700-00AA0060263B}' : u'AddIns',
	'{91493461-5A91-11CF-8700-00AA0060263B}' : u'AddIn',
	'{91493462-5A91-11CF-8700-00AA0060263B}' : u'Presentations',
	'{91493464-5A91-11CF-8700-00AA0060263B}' : u'Hyperlinks',
	'{91493465-5A91-11CF-8700-00AA0060263B}' : u'Hyperlink',
	'{91493466-5A91-11CF-8700-00AA0060263B}' : u'PageSetup',
	'{91493467-5A91-11CF-8700-00AA0060263B}' : u'Fonts',
	'{91493468-5A91-11CF-8700-00AA0060263B}' : u'ExtraColors',
	'{91493469-5A91-11CF-8700-00AA0060263B}' : u'Slides',
	'{9149346A-5A91-11CF-8700-00AA0060263B}' : u'_Slide',
	'{9149346B-5A91-11CF-8700-00AA0060263B}' : u'SlideRange',
	'{9149346C-5A91-11CF-8700-00AA0060263B}' : u'_Master',
	'{9149346E-5A91-11CF-8700-00AA0060263B}' : u'ColorSchemes',
	'{9149346F-5A91-11CF-8700-00AA0060263B}' : u'ColorScheme',
	'{91493470-5A91-11CF-8700-00AA0060263B}' : u'RGBColor',
	'{91493471-5A91-11CF-8700-00AA0060263B}' : u'SlideShowTransition',
	'{91493472-5A91-11CF-8700-00AA0060263B}' : u'SoundEffect',
	'{91493473-5A91-11CF-8700-00AA0060263B}' : u'SoundFormat',
	'{91493474-5A91-11CF-8700-00AA0060263B}' : u'HeadersFooters',
	'{91493475-5A91-11CF-8700-00AA0060263B}' : u'Shapes',
	'{91493476-5A91-11CF-8700-00AA0060263B}' : u'Placeholders',
	'{91493477-5A91-11CF-8700-00AA0060263B}' : u'PlaceholderFormat',
	'{91493478-5A91-11CF-8700-00AA0060263B}' : u'FreeformBuilder',
	'{91493479-5A91-11CF-8700-00AA0060263B}' : u'Shape',
	'{9149347A-5A91-11CF-8700-00AA0060263B}' : u'ShapeRange',
	'{9149347B-5A91-11CF-8700-00AA0060263B}' : u'GroupShapes',
	'{9149347C-5A91-11CF-8700-00AA0060263B}' : u'Adjustments',
	'{9149347D-5A91-11CF-8700-00AA0060263B}' : u'PictureFormat',
	'{9149347E-5A91-11CF-8700-00AA0060263B}' : u'FillFormat',
	'{9149347F-5A91-11CF-8700-00AA0060263B}' : u'LineFormat',
	'{91493480-5A91-11CF-8700-00AA0060263B}' : u'ShadowFormat',
	'{91493481-5A91-11CF-8700-00AA0060263B}' : u'ConnectorFormat',
	'{91493482-5A91-11CF-8700-00AA0060263B}' : u'TextEffectFormat',
	'{91493483-5A91-11CF-8700-00AA0060263B}' : u'ThreeDFormat',
	'{91493484-5A91-11CF-8700-00AA0060263B}' : u'TextFrame',
	'{91493485-5A91-11CF-8700-00AA0060263B}' : u'CalloutFormat',
	'{91493486-5A91-11CF-8700-00AA0060263B}' : u'ShapeNodes',
	'{91493487-5A91-11CF-8700-00AA0060263B}' : u'ShapeNode',
	'{91493488-5A91-11CF-8700-00AA0060263B}' : u'OLEFormat',
	'{91493489-5A91-11CF-8700-00AA0060263B}' : u'LinkFormat',
	'{9149348A-5A91-11CF-8700-00AA0060263B}' : u'ObjectVerbs',
	'{9149348B-5A91-11CF-8700-00AA0060263B}' : u'AnimationSettings',
	'{9149348C-5A91-11CF-8700-00AA0060263B}' : u'ActionSettings',
	'{9149348D-5A91-11CF-8700-00AA0060263B}' : u'ActionSetting',
	'{9149348E-5A91-11CF-8700-00AA0060263B}' : u'PlaySettings',
	'{9149348F-5A91-11CF-8700-00AA0060263B}' : u'TextRange',
	'{91493490-5A91-11CF-8700-00AA0060263B}' : u'Ruler',
	'{91493491-5A91-11CF-8700-00AA0060263B}' : u'RulerLevels',
	'{91493492-5A91-11CF-8700-00AA0060263B}' : u'RulerLevel',
	'{91493493-5A91-11CF-8700-00AA0060263B}' : u'TabStops',
	'{91493494-5A91-11CF-8700-00AA0060263B}' : u'TabStop',
	'{91493495-5A91-11CF-8700-00AA0060263B}' : u'Font',
	'{91493496-5A91-11CF-8700-00AA0060263B}' : u'ParagraphFormat',
	'{91493497-5A91-11CF-8700-00AA0060263B}' : u'BulletFormat',
	'{91493498-5A91-11CF-8700-00AA0060263B}' : u'TextStyles',
	'{91493499-5A91-11CF-8700-00AA0060263B}' : u'TextStyle',
	'{9149349A-5A91-11CF-8700-00AA0060263B}' : u'TextStyleLevels',
	'{9149349B-5A91-11CF-8700-00AA0060263B}' : u'TextStyleLevel',
	'{9149349C-5A91-11CF-8700-00AA0060263B}' : u'HeaderFooter',
	'{9149349D-5A91-11CF-8700-00AA0060263B}' : u'_Presentation',
	'{914934B9-5A91-11CF-8700-00AA0060263B}' : u'Tags',
	'{914934C0-5A91-11CF-8700-00AA0060263B}' : u'OCXExtender',
	'{914934C1-5A91-11CF-8700-00AA0060263B}' : u'OCXExtenderEvents',
	'{914934C2-5A91-11CF-8700-00AA0060263B}' : u'EApplication',
	'{914934C3-5A91-11CF-8700-00AA0060263B}' : u'Table',
	'{914934C4-5A91-11CF-8700-00AA0060263B}' : u'Columns',
	'{914934C5-5A91-11CF-8700-00AA0060263B}' : u'Column',
	'{914934C6-5A91-11CF-8700-00AA0060263B}' : u'Rows',
	'{914934C7-5A91-11CF-8700-00AA0060263B}' : u'Row',
	'{914934C8-5A91-11CF-8700-00AA0060263B}' : u'CellRange',
	'{914934C9-5A91-11CF-8700-00AA0060263B}' : u'Cell',
	'{914934CA-5A91-11CF-8700-00AA0060263B}' : u'Borders',
	'{914934CB-5A91-11CF-8700-00AA0060263B}' : u'Panes',
	'{914934CC-5A91-11CF-8700-00AA0060263B}' : u'Pane',
	'{914934CD-5A91-11CF-8700-00AA0060263B}' : u'DefaultWebOptions',
	'{914934CE-5A91-11CF-8700-00AA0060263B}' : u'WebOptions',
	'{914934CF-5A91-11CF-8700-00AA0060263B}' : u'PublishObjects',
	'{914934D0-5A91-11CF-8700-00AA0060263B}' : u'PublishObject',
	'{914934D3-5A91-11CF-8700-00AA0060263B}' : u'_PowerRex',
	'{914934D4-5A91-11CF-8700-00AA0060263B}' : u'Comments',
	'{914934D5-5A91-11CF-8700-00AA0060263B}' : u'Comment',
	'{914934D6-5A91-11CF-8700-00AA0060263B}' : u'Designs',
	'{914934D7-5A91-11CF-8700-00AA0060263B}' : u'Design',
	'{914934D8-5A91-11CF-8700-00AA0060263B}' : u'DiagramNode',
	'{914934D9-5A91-11CF-8700-00AA0060263B}' : u'DiagramNodeChildren',
	'{914934DA-5A91-11CF-8700-00AA0060263B}' : u'DiagramNodes',
	'{914934DB-5A91-11CF-8700-00AA0060263B}' : u'Diagram',
	'{914934DC-5A91-11CF-8700-00AA0060263B}' : u'TimeLine',
	'{914934DD-5A91-11CF-8700-00AA0060263B}' : u'Sequences',
	'{914934DE-5A91-11CF-8700-00AA0060263B}' : u'Sequence',
	'{914934DF-5A91-11CF-8700-00AA0060263B}' : u'Effect',
	'{914934E0-5A91-11CF-8700-00AA0060263B}' : u'Timing',
	'{914934E1-5A91-11CF-8700-00AA0060263B}' : u'EffectParameters',
	'{914934E2-5A91-11CF-8700-00AA0060263B}' : u'EffectInformation',
	'{914934E3-5A91-11CF-8700-00AA0060263B}' : u'AnimationBehaviors',
	'{914934E4-5A91-11CF-8700-00AA0060263B}' : u'AnimationBehavior',
	'{914934E5-5A91-11CF-8700-00AA0060263B}' : u'MotionEffect',
	'{914934E6-5A91-11CF-8700-00AA0060263B}' : u'ColorEffect',
	'{914934E7-5A91-11CF-8700-00AA0060263B}' : u'ScaleEffect',
}
VTablesToClassMap = {}
VTablesToPackageMap = {
	'{914934E8-5A91-11CF-8700-00AA0060263B}' : 'RotationEffect',
	'{914934E9-5A91-11CF-8700-00AA0060263B}' : 'PropertyEffect',
	'{914934EA-5A91-11CF-8700-00AA0060263B}' : 'AnimationPoints',
	'{914934EB-5A91-11CF-8700-00AA0060263B}' : 'AnimationPoint',
	'{914934EC-5A91-11CF-8700-00AA0060263B}' : 'CanvasShapes',
	'{914934ED-5A91-11CF-8700-00AA0060263B}' : 'AutoCorrect',
	'{914934EE-5A91-11CF-8700-00AA0060263B}' : 'Options',
	'{914934EF-5A91-11CF-8700-00AA0060263B}' : 'CommandEffect',
	'{914934F0-5A91-11CF-8700-00AA0060263B}' : 'FilterEffect',
	'{914934F1-5A91-11CF-8700-00AA0060263B}' : 'SetEffect',
	'{914934F2-5A91-11CF-8700-00AA0060263B}' : 'CustomLayouts',
	'{914934F3-5A91-11CF-8700-00AA0060263B}' : 'CustomLayout',
	'{914934F5-5A91-11CF-8700-00AA0060263B}' : 'TableStyle',
	'{914934F6-5A91-11CF-8700-00AA0060263B}' : 'CustomerData',
	'{914934F7-5A91-11CF-8700-00AA0060263B}' : 'Research',
	'{914934F8-5A91-11CF-8700-00AA0060263B}' : 'TableBackground',
	'{914934F9-5A91-11CF-8700-00AA0060263B}' : 'TextFrame2',
	'{91493442-5A91-11CF-8700-00AA0060263B}' : '_Application',
	'{91493450-5A91-11CF-8700-00AA0060263B}' : 'Collection',
	'{91493451-5A91-11CF-8700-00AA0060263B}' : '_Global',
	'{91493452-5A91-11CF-8700-00AA0060263B}' : 'ColorFormat',
	'{91493453-5A91-11CF-8700-00AA0060263B}' : 'SlideShowWindow',
	'{91493454-5A91-11CF-8700-00AA0060263B}' : 'Selection',
	'{91493455-5A91-11CF-8700-00AA0060263B}' : 'DocumentWindows',
	'{91493456-5A91-11CF-8700-00AA0060263B}' : 'SlideShowWindows',
	'{91493457-5A91-11CF-8700-00AA0060263B}' : 'DocumentWindow',
	'{91493458-5A91-11CF-8700-00AA0060263B}' : 'View',
	'{91493459-5A91-11CF-8700-00AA0060263B}' : 'SlideShowView',
	'{9149345A-5A91-11CF-8700-00AA0060263B}' : 'SlideShowSettings',
	'{9149345B-5A91-11CF-8700-00AA0060263B}' : 'NamedSlideShows',
	'{9149345C-5A91-11CF-8700-00AA0060263B}' : 'NamedSlideShow',
	'{9149345D-5A91-11CF-8700-00AA0060263B}' : 'PrintOptions',
	'{9149345E-5A91-11CF-8700-00AA0060263B}' : 'PrintRanges',
	'{9149345F-5A91-11CF-8700-00AA0060263B}' : 'PrintRange',
	'{91493460-5A91-11CF-8700-00AA0060263B}' : 'AddIns',
	'{91493461-5A91-11CF-8700-00AA0060263B}' : 'AddIn',
	'{91493462-5A91-11CF-8700-00AA0060263B}' : 'Presentations',
	'{91493463-5A91-11CF-8700-00AA0060263B}' : 'PresEvents',
	'{91493464-5A91-11CF-8700-00AA0060263B}' : 'Hyperlinks',
	'{91493465-5A91-11CF-8700-00AA0060263B}' : 'Hyperlink',
	'{91493466-5A91-11CF-8700-00AA0060263B}' : 'PageSetup',
	'{91493467-5A91-11CF-8700-00AA0060263B}' : 'Fonts',
	'{91493468-5A91-11CF-8700-00AA0060263B}' : 'ExtraColors',
	'{91493469-5A91-11CF-8700-00AA0060263B}' : 'Slides',
	'{9149346A-5A91-11CF-8700-00AA0060263B}' : '_Slide',
	'{9149346B-5A91-11CF-8700-00AA0060263B}' : 'SlideRange',
	'{9149346C-5A91-11CF-8700-00AA0060263B}' : '_Master',
	'{9149346D-5A91-11CF-8700-00AA0060263B}' : 'SldEvents',
	'{9149346E-5A91-11CF-8700-00AA0060263B}' : 'ColorSchemes',
	'{9149346F-5A91-11CF-8700-00AA0060263B}' : 'ColorScheme',
	'{91493470-5A91-11CF-8700-00AA0060263B}' : 'RGBColor',
	'{91493471-5A91-11CF-8700-00AA0060263B}' : 'SlideShowTransition',
	'{91493472-5A91-11CF-8700-00AA0060263B}' : 'SoundEffect',
	'{91493473-5A91-11CF-8700-00AA0060263B}' : 'SoundFormat',
	'{91493474-5A91-11CF-8700-00AA0060263B}' : 'HeadersFooters',
	'{91493475-5A91-11CF-8700-00AA0060263B}' : 'Shapes',
	'{91493476-5A91-11CF-8700-00AA0060263B}' : 'Placeholders',
	'{91493477-5A91-11CF-8700-00AA0060263B}' : 'PlaceholderFormat',
	'{91493478-5A91-11CF-8700-00AA0060263B}' : 'FreeformBuilder',
	'{91493479-5A91-11CF-8700-00AA0060263B}' : 'Shape',
	'{9149347A-5A91-11CF-8700-00AA0060263B}' : 'ShapeRange',
	'{9149347B-5A91-11CF-8700-00AA0060263B}' : 'GroupShapes',
	'{9149347C-5A91-11CF-8700-00AA0060263B}' : 'Adjustments',
	'{9149347D-5A91-11CF-8700-00AA0060263B}' : 'PictureFormat',
	'{9149347E-5A91-11CF-8700-00AA0060263B}' : 'FillFormat',
	'{9149347F-5A91-11CF-8700-00AA0060263B}' : 'LineFormat',
	'{91493480-5A91-11CF-8700-00AA0060263B}' : 'ShadowFormat',
	'{91493481-5A91-11CF-8700-00AA0060263B}' : 'ConnectorFormat',
	'{91493482-5A91-11CF-8700-00AA0060263B}' : 'TextEffectFormat',
	'{91493483-5A91-11CF-8700-00AA0060263B}' : 'ThreeDFormat',
	'{91493484-5A91-11CF-8700-00AA0060263B}' : 'TextFrame',
	'{91493485-5A91-11CF-8700-00AA0060263B}' : 'CalloutFormat',
	'{91493486-5A91-11CF-8700-00AA0060263B}' : 'ShapeNodes',
	'{91493487-5A91-11CF-8700-00AA0060263B}' : 'ShapeNode',
	'{91493488-5A91-11CF-8700-00AA0060263B}' : 'OLEFormat',
	'{91493489-5A91-11CF-8700-00AA0060263B}' : 'LinkFormat',
	'{9149348A-5A91-11CF-8700-00AA0060263B}' : 'ObjectVerbs',
	'{9149348B-5A91-11CF-8700-00AA0060263B}' : 'AnimationSettings',
	'{9149348C-5A91-11CF-8700-00AA0060263B}' : 'ActionSettings',
	'{9149348D-5A91-11CF-8700-00AA0060263B}' : 'ActionSetting',
	'{9149348E-5A91-11CF-8700-00AA0060263B}' : 'PlaySettings',
	'{9149348F-5A91-11CF-8700-00AA0060263B}' : 'TextRange',
	'{91493490-5A91-11CF-8700-00AA0060263B}' : 'Ruler',
	'{91493491-5A91-11CF-8700-00AA0060263B}' : 'RulerLevels',
	'{91493492-5A91-11CF-8700-00AA0060263B}' : 'RulerLevel',
	'{91493493-5A91-11CF-8700-00AA0060263B}' : 'TabStops',
	'{91493494-5A91-11CF-8700-00AA0060263B}' : 'TabStop',
	'{91493495-5A91-11CF-8700-00AA0060263B}' : 'Font',
	'{91493496-5A91-11CF-8700-00AA0060263B}' : 'ParagraphFormat',
	'{91493497-5A91-11CF-8700-00AA0060263B}' : 'BulletFormat',
	'{91493498-5A91-11CF-8700-00AA0060263B}' : 'TextStyles',
	'{91493499-5A91-11CF-8700-00AA0060263B}' : 'TextStyle',
	'{9149349A-5A91-11CF-8700-00AA0060263B}' : 'TextStyleLevels',
	'{9149349B-5A91-11CF-8700-00AA0060263B}' : 'TextStyleLevel',
	'{9149349C-5A91-11CF-8700-00AA0060263B}' : 'HeaderFooter',
	'{9149349D-5A91-11CF-8700-00AA0060263B}' : '_Presentation',
	'{914934B9-5A91-11CF-8700-00AA0060263B}' : 'Tags',
	'{914934BE-5A91-11CF-8700-00AA0060263B}' : 'MouseTracker',
	'{914934BF-5A91-11CF-8700-00AA0060263B}' : 'MouseDownHandler',
	'{914934C0-5A91-11CF-8700-00AA0060263B}' : 'OCXExtender',
	'{914934C3-5A91-11CF-8700-00AA0060263B}' : 'Table',
	'{914934C4-5A91-11CF-8700-00AA0060263B}' : 'Columns',
	'{914934C5-5A91-11CF-8700-00AA0060263B}' : 'Column',
	'{914934C6-5A91-11CF-8700-00AA0060263B}' : 'Rows',
	'{914934C7-5A91-11CF-8700-00AA0060263B}' : 'Row',
	'{914934C8-5A91-11CF-8700-00AA0060263B}' : 'CellRange',
	'{914934C9-5A91-11CF-8700-00AA0060263B}' : 'Cell',
	'{914934CA-5A91-11CF-8700-00AA0060263B}' : 'Borders',
	'{914934CB-5A91-11CF-8700-00AA0060263B}' : 'Panes',
	'{914934CC-5A91-11CF-8700-00AA0060263B}' : 'Pane',
	'{914934CD-5A91-11CF-8700-00AA0060263B}' : 'DefaultWebOptions',
	'{914934CE-5A91-11CF-8700-00AA0060263B}' : 'WebOptions',
	'{914934CF-5A91-11CF-8700-00AA0060263B}' : 'PublishObjects',
	'{914934D0-5A91-11CF-8700-00AA0060263B}' : 'PublishObject',
	'{914934D2-5A91-11CF-8700-00AA0060263B}' : 'MasterEvents',
	'{914934D3-5A91-11CF-8700-00AA0060263B}' : '_PowerRex',
	'{914934D4-5A91-11CF-8700-00AA0060263B}' : 'Comments',
	'{914934D5-5A91-11CF-8700-00AA0060263B}' : 'Comment',
	'{914934D6-5A91-11CF-8700-00AA0060263B}' : 'Designs',
	'{914934D7-5A91-11CF-8700-00AA0060263B}' : 'Design',
	'{914934D8-5A91-11CF-8700-00AA0060263B}' : 'DiagramNode',
	'{914934D9-5A91-11CF-8700-00AA0060263B}' : 'DiagramNodeChildren',
	'{914934DA-5A91-11CF-8700-00AA0060263B}' : 'DiagramNodes',
	'{914934DB-5A91-11CF-8700-00AA0060263B}' : 'Diagram',
	'{914934DC-5A91-11CF-8700-00AA0060263B}' : 'TimeLine',
	'{914934DD-5A91-11CF-8700-00AA0060263B}' : 'Sequences',
	'{914934DE-5A91-11CF-8700-00AA0060263B}' : 'Sequence',
	'{914934DF-5A91-11CF-8700-00AA0060263B}' : 'Effect',
	'{914934E0-5A91-11CF-8700-00AA0060263B}' : 'Timing',
	'{914934E1-5A91-11CF-8700-00AA0060263B}' : 'EffectParameters',
	'{914934E2-5A91-11CF-8700-00AA0060263B}' : 'EffectInformation',
	'{914934E3-5A91-11CF-8700-00AA0060263B}' : 'AnimationBehaviors',
	'{914934E4-5A91-11CF-8700-00AA0060263B}' : 'AnimationBehavior',
	'{914934E5-5A91-11CF-8700-00AA0060263B}' : 'MotionEffect',
	'{914934E6-5A91-11CF-8700-00AA0060263B}' : 'ColorEffect',
	'{914934E7-5A91-11CF-8700-00AA0060263B}' : 'ScaleEffect',
}


NamesToIIDMap = {
	'CanvasShapes' : '{914934EC-5A91-11CF-8700-00AA0060263B}',
	'Research' : '{914934F7-5A91-11CF-8700-00AA0060263B}',
	'RulerLevel' : '{91493492-5A91-11CF-8700-00AA0060263B}',
	'MasterEvents' : '{914934D2-5A91-11CF-8700-00AA0060263B}',
	'ThreeDFormat' : '{91493483-5A91-11CF-8700-00AA0060263B}',
	'PresEvents' : '{91493463-5A91-11CF-8700-00AA0060263B}',
	'ColorScheme' : '{9149346F-5A91-11CF-8700-00AA0060263B}',
	'Ruler' : '{91493490-5A91-11CF-8700-00AA0060263B}',
	'ObjectVerbs' : '{9149348A-5A91-11CF-8700-00AA0060263B}',
	'CustomLayout' : '{914934F3-5A91-11CF-8700-00AA0060263B}',
	'ActionSettings' : '{9149348C-5A91-11CF-8700-00AA0060263B}',
	'TabStops' : '{91493493-5A91-11CF-8700-00AA0060263B}',
	'AddIn' : '{91493461-5A91-11CF-8700-00AA0060263B}',
	'DefaultWebOptions' : '{914934CD-5A91-11CF-8700-00AA0060263B}',
	'PublishObject' : '{914934D0-5A91-11CF-8700-00AA0060263B}',
	'CalloutFormat' : '{91493485-5A91-11CF-8700-00AA0060263B}',
	'Comments' : '{914934D4-5A91-11CF-8700-00AA0060263B}',
	'PictureFormat' : '{9149347D-5A91-11CF-8700-00AA0060263B}',
	'MouseDownHandler' : '{914934BF-5A91-11CF-8700-00AA0060263B}',
	'_PowerRex' : '{914934D3-5A91-11CF-8700-00AA0060263B}',
	'ShapeNodes' : '{91493486-5A91-11CF-8700-00AA0060263B}',
	'_Global' : '{91493451-5A91-11CF-8700-00AA0060263B}',
	'Pane' : '{914934CC-5A91-11CF-8700-00AA0060263B}',
	'MotionEffect' : '{914934E5-5A91-11CF-8700-00AA0060263B}',
	'Font' : '{91493495-5A91-11CF-8700-00AA0060263B}',
	'TextStyle' : '{91493499-5A91-11CF-8700-00AA0060263B}',
	'_Application' : '{91493442-5A91-11CF-8700-00AA0060263B}',
	'TableBackground' : '{914934F8-5A91-11CF-8700-00AA0060263B}',
	'Hyperlinks' : '{91493464-5A91-11CF-8700-00AA0060263B}',
	'Placeholders' : '{91493476-5A91-11CF-8700-00AA0060263B}',
	'_Slide' : '{9149346A-5A91-11CF-8700-00AA0060263B}',
	'AnimationPoint' : '{914934EB-5A91-11CF-8700-00AA0060263B}',
	'WebOptions' : '{914934CE-5A91-11CF-8700-00AA0060263B}',
	'LinkFormat' : '{91493489-5A91-11CF-8700-00AA0060263B}',
	'EffectParameters' : '{914934E1-5A91-11CF-8700-00AA0060263B}',
	'ShapeRange' : '{9149347A-5A91-11CF-8700-00AA0060263B}',
	'ColorEffect' : '{914934E6-5A91-11CF-8700-00AA0060263B}',
	'FillFormat' : '{9149347E-5A91-11CF-8700-00AA0060263B}',
	'AddIns' : '{91493460-5A91-11CF-8700-00AA0060263B}',
	'NamedSlideShows' : '{9149345B-5A91-11CF-8700-00AA0060263B}',
	'PrintRanges' : '{9149345E-5A91-11CF-8700-00AA0060263B}',
	'PublishObjects' : '{914934CF-5A91-11CF-8700-00AA0060263B}',
	'ShadowFormat' : '{91493480-5A91-11CF-8700-00AA0060263B}',
	'TimeLine' : '{914934DC-5A91-11CF-8700-00AA0060263B}',
	'DocumentWindow' : '{91493457-5A91-11CF-8700-00AA0060263B}',
	'Presentations' : '{91493462-5A91-11CF-8700-00AA0060263B}',
	'Shapes' : '{91493475-5A91-11CF-8700-00AA0060263B}',
	'PlaySettings' : '{9149348E-5A91-11CF-8700-00AA0060263B}',
	'OLEFormat' : '{91493488-5A91-11CF-8700-00AA0060263B}',
	'SoundEffect' : '{91493472-5A91-11CF-8700-00AA0060263B}',
	'NamedSlideShow' : '{9149345C-5A91-11CF-8700-00AA0060263B}',
	'AnimationSettings' : '{9149348B-5A91-11CF-8700-00AA0060263B}',
	'FreeformBuilder' : '{91493478-5A91-11CF-8700-00AA0060263B}',
	'Options' : '{914934EE-5A91-11CF-8700-00AA0060263B}',
	'Comment' : '{914934D5-5A91-11CF-8700-00AA0060263B}',
	'Selection' : '{91493454-5A91-11CF-8700-00AA0060263B}',
	'TableStyle' : '{914934F5-5A91-11CF-8700-00AA0060263B}',
	'Tags' : '{914934B9-5A91-11CF-8700-00AA0060263B}',
	'BulletFormat' : '{91493497-5A91-11CF-8700-00AA0060263B}',
	'EApplication' : '{914934C2-5A91-11CF-8700-00AA0060263B}',
	'Shape' : '{91493479-5A91-11CF-8700-00AA0060263B}',
	'Design' : '{914934D7-5A91-11CF-8700-00AA0060263B}',
	'LineFormat' : '{9149347F-5A91-11CF-8700-00AA0060263B}',
	'Row' : '{914934C7-5A91-11CF-8700-00AA0060263B}',
	'SlideShowSettings' : '{9149345A-5A91-11CF-8700-00AA0060263B}',
	'PrintOptions' : '{9149345D-5A91-11CF-8700-00AA0060263B}',
	'PrintRange' : '{9149345F-5A91-11CF-8700-00AA0060263B}',
	'CustomLayouts' : '{914934F2-5A91-11CF-8700-00AA0060263B}',
	'Cell' : '{914934C9-5A91-11CF-8700-00AA0060263B}',
	'HeadersFooters' : '{91493474-5A91-11CF-8700-00AA0060263B}',
	'CellRange' : '{914934C8-5A91-11CF-8700-00AA0060263B}',
	'SlideShowWindows' : '{91493456-5A91-11CF-8700-00AA0060263B}',
	'TextStyleLevel' : '{9149349B-5A91-11CF-8700-00AA0060263B}',
	'Collection' : '{91493450-5A91-11CF-8700-00AA0060263B}',
	'Borders' : '{914934CA-5A91-11CF-8700-00AA0060263B}',
	'ColorFormat' : '{91493452-5A91-11CF-8700-00AA0060263B}',
	'TextRange' : '{9149348F-5A91-11CF-8700-00AA0060263B}',
	'Columns' : '{914934C4-5A91-11CF-8700-00AA0060263B}',
	'ExtraColors' : '{91493468-5A91-11CF-8700-00AA0060263B}',
	'Designs' : '{914934D6-5A91-11CF-8700-00AA0060263B}',
	'DiagramNode' : '{914934D8-5A91-11CF-8700-00AA0060263B}',
	'RotationEffect' : '{914934E8-5A91-11CF-8700-00AA0060263B}',
	'AnimationBehavior' : '{914934E4-5A91-11CF-8700-00AA0060263B}',
	'CommandEffect' : '{914934EF-5A91-11CF-8700-00AA0060263B}',
	'PropertyEffect' : '{914934E9-5A91-11CF-8700-00AA0060263B}',
	'ParagraphFormat' : '{91493496-5A91-11CF-8700-00AA0060263B}',
	'DiagramNodes' : '{914934DA-5A91-11CF-8700-00AA0060263B}',
	'ConnectorFormat' : '{91493481-5A91-11CF-8700-00AA0060263B}',
	'EffectInformation' : '{914934E2-5A91-11CF-8700-00AA0060263B}',
	'HeaderFooter' : '{9149349C-5A91-11CF-8700-00AA0060263B}',
	'SlideShowWindow' : '{91493453-5A91-11CF-8700-00AA0060263B}',
	'SetEffect' : '{914934F1-5A91-11CF-8700-00AA0060263B}',
	'_Presentation' : '{9149349D-5A91-11CF-8700-00AA0060263B}',
	'Column' : '{914934C5-5A91-11CF-8700-00AA0060263B}',
	'Fonts' : '{91493467-5A91-11CF-8700-00AA0060263B}',
	'SlideRange' : '{9149346B-5A91-11CF-8700-00AA0060263B}',
	'DiagramNodeChildren' : '{914934D9-5A91-11CF-8700-00AA0060263B}',
	'ActionSetting' : '{9149348D-5A91-11CF-8700-00AA0060263B}',
	'OCXExtender' : '{914934C0-5A91-11CF-8700-00AA0060263B}',
	'Effect' : '{914934DF-5A91-11CF-8700-00AA0060263B}',
	'ShapeNode' : '{91493487-5A91-11CF-8700-00AA0060263B}',
	'RGBColor' : '{91493470-5A91-11CF-8700-00AA0060263B}',
	'DocumentWindows' : '{91493455-5A91-11CF-8700-00AA0060263B}',
	'PlaceholderFormat' : '{91493477-5A91-11CF-8700-00AA0060263B}',
	'FilterEffect' : '{914934F0-5A91-11CF-8700-00AA0060263B}',
	'SldEvents' : '{9149346D-5A91-11CF-8700-00AA0060263B}',
	'Rows' : '{914934C6-5A91-11CF-8700-00AA0060263B}',
	'TextFrame2' : '{914934F9-5A91-11CF-8700-00AA0060263B}',
	'ScaleEffect' : '{914934E7-5A91-11CF-8700-00AA0060263B}',
	'TextFrame' : '{91493484-5A91-11CF-8700-00AA0060263B}',
	'Panes' : '{914934CB-5A91-11CF-8700-00AA0060263B}',
	'RulerLevels' : '{91493491-5A91-11CF-8700-00AA0060263B}',
	'MouseTracker' : '{914934BE-5A91-11CF-8700-00AA0060263B}',
	'OCXExtenderEvents' : '{914934C1-5A91-11CF-8700-00AA0060263B}',
	'TabStop' : '{91493494-5A91-11CF-8700-00AA0060263B}',
	'CustomerData' : '{914934F6-5A91-11CF-8700-00AA0060263B}',
	'Slides' : '{91493469-5A91-11CF-8700-00AA0060263B}',
	'TextStyleLevels' : '{9149349A-5A91-11CF-8700-00AA0060263B}',
	'Timing' : '{914934E0-5A91-11CF-8700-00AA0060263B}',
	'SoundFormat' : '{91493473-5A91-11CF-8700-00AA0060263B}',
	'Hyperlink' : '{91493465-5A91-11CF-8700-00AA0060263B}',
	'AnimationBehaviors' : '{914934E3-5A91-11CF-8700-00AA0060263B}',
	'SlideShowView' : '{91493459-5A91-11CF-8700-00AA0060263B}',
	'TextEffectFormat' : '{91493482-5A91-11CF-8700-00AA0060263B}',
	'AnimationPoints' : '{914934EA-5A91-11CF-8700-00AA0060263B}',
	'_Master' : '{9149346C-5A91-11CF-8700-00AA0060263B}',
	'View' : '{91493458-5A91-11CF-8700-00AA0060263B}',
	'TextStyles' : '{91493498-5A91-11CF-8700-00AA0060263B}',
	'Diagram' : '{914934DB-5A91-11CF-8700-00AA0060263B}',
	'Sequence' : '{914934DE-5A91-11CF-8700-00AA0060263B}',
	'AutoCorrect' : '{914934ED-5A91-11CF-8700-00AA0060263B}',
	'Adjustments' : '{9149347C-5A91-11CF-8700-00AA0060263B}',
	'SlideShowTransition' : '{91493471-5A91-11CF-8700-00AA0060263B}',
	'GroupShapes' : '{9149347B-5A91-11CF-8700-00AA0060263B}',
	'Sequences' : '{914934DD-5A91-11CF-8700-00AA0060263B}',
	'Table' : '{914934C3-5A91-11CF-8700-00AA0060263B}',
	'ColorSchemes' : '{9149346E-5A91-11CF-8700-00AA0060263B}',
	'PageSetup' : '{91493466-5A91-11CF-8700-00AA0060263B}',
}

win32com.client.constants.__dicts__.append(constants.__dict__)

