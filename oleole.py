# -*- coding: utf-8 -*- 

import sys
import os
from oletools.olevba import VBA_Parser, TYPE_OLE, TYPE_OpenXML, TYPE_Word2003_XML, TYPE_MHTML, detect_autoexec, VBA_Scanner

#PYTHON 2,3버전 모두 호환 
if sys.version_info[0]<= 2:
      #python 2
      from oletools.olevba import VBA_Parser
else:
     #python 3
	  from oletools.olevba3 import VBA_Parser

# VBA 퍼져 온 
vbaparser = VBA_Parser(sys.argv[1])
  # 매크로 검사
if not vbaparser.detect_vba_macros():
 print ('error', "macro not detected ")
 #경고모듈 넣는 위치 
 sys.exit(1)

print ('info', "macro detected ! ")

for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
  print ("- "*79)
  print ("filename    :", filename)
  print ("OLE stream  :", stream_path)
  print ("VBA filename:", vba_filename)
  print ("- "*39)
  print ("vba_code",vba_code)

     
vba_scanner = VBA_Scanner(vba_code)
results = vba_scanner.scan(include_decoded_strings=True)
for kw_type, keyword, description in results:
 print ("type=%s - keyword=%s - description=%s" % (kw_type, keyword, description))
     
autoexec_keywords = detect_autoexec(vba_code)
if autoexec_keywords:
 print("Auto-executable Macro keyword is detected. :")
for keyword, description in autoexec_keywords:
   print("%s: %s' % (keyword, description)")
else:
 print ("Auto-executable Macro keywords is not detected.")

results = vbaparser.analyze_macros()
for kw_type, keyword, description in results:
 print ("type=%s - keyword=%s - description=%s" % (kw_type, keyword, description))
 print ("AutoExec keywords: %d" % vbaparser.nb_autoexec)
 print ("Suspicious keywords: %d" % vbaparser.nb_suspicious)
 print ("IOCs: %d" % vbaparser.nb_iocs)
 print ("Hex obfuscated strings: %d" % vbaparser.nb_hexstrings)
 print ("Base64 obfuscated strings: %d" % vbaparser.nb_base64strings)
 print ("Dridex obfuscated strings: %d" % vbaparser.nb_dridexstrings)
 print ("VBA obfuscated strings: %d" % vbaparser.nb_vbastrings)

  # 퍼져 닫기 
vbaparser.close()
