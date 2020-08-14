# olevba_tool
olevba base _ pauser tool

기본 설치 
pip install python-ooxml
pip install oletools 


Olevba tools를 기반으로 하는 pauser tool

1. oleole.py 

 oletools.olevba에서 
 VBA_Parser, TYPE_OLE, TYPE_OpenXML, TYPE_Word2003_XML, TYPE_MHTML, detect_autoexec, VBA_Scanner를 import하여 사용합니다.
 
 
  입력 : python oleole.py 파일명
 
  출력 : 
 
![주석 2020-08-14 2203572](https://user-images.githubusercontent.com/67878157/90252745-fa017f80-de7a-11ea-829c-86f654115eae.png)
![주석 2020-08-14 220357](https://user-images.githubusercontent.com/67878157/90252695-e524ec00-de7a-11ea-9565-c65c0b769099.png)

 
 
 
 해당 툴은 파이선 2,3버전 모두와 호환됩니다. 
  
    프로그램 실행순서
    1. 명령줄 인수로 받은 파일을 VBA_Parse를 통해 매크로 검사 
    2. 해당 파일을 VBA_Scanner를 활용해 코드를 추출함 
    3. 어떤 유형의 난독화인지 찾아내고, 이를 화면에 출력  
    


2. ooxmlpaser.py
Python-OOXML은 Office Open XML 파일을 구문 분석하기위한 Python 라이브러리입니다. 
현재 HTML 만 출력 형식으로 지원합니다. 출력의 쉬운 사용자 정의에 중점을 둡니다. 
라이브러리에는 문서를 별도의 장으로 분할 할 수있는 가져 오기 도구가 함께 제공됩니다. Word 스타일을 사용하는 문서와 사용되지 않는 문서 모두에서 작동합니다.

