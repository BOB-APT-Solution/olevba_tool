# olevba_tool
olevba base _ pauser tool

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
    
