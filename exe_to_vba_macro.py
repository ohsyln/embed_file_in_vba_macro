f=open('input.txt','rb')
s=f.read()
f.close()

bytes_per_line = 30
lines_per_function = 50
bytes_per_function = bytes_per_line * lines_per_function
functions_per_line = 15

out_msg = ''
very_first_line = 1
first_line = 1
print(len(s))
for i in range(0,len(s)):
  portion = i//bytes_per_function
  if i % bytes_per_function == 0:  
    out_msg += 'function a{} as string\n'.format(portion)
    first_line = 1
  if i % bytes_per_line == 0:
    if very_first_line:
      out_msg+='a{}="{}'.format(portion,s[i])
      very_first_line = 0
      first_line = 0
    elif first_line:
      out_msg+='a{}=",{}'.format(portion,s[i])
      first_line = 0
    else:
      out_msg+='a{p}=a{p}+",{c}'.format(p=portion,c=s[i])
  else:
    out_msg += ',{}'.format(s[i])
  if i % bytes_per_line == bytes_per_line - 1: # last
    out_msg += '"\n'
  if i % bytes_per_function == bytes_per_function - 1:
    out_msg += '\nend function\n'
    first_line = 1
    
out_msg += '"\nend function\n\n'
out_msg += 'sub autoopen()\n\n'
out_msg += 'dim arr\n'
out_msg += 'arr = ""\n'
for k in range(0,portion+1):
  if k % functions_per_line == 0:
    out_msg += 'arr = arr + a{}()'.format(k)
  else:
    out_msg += '+ a{}()'.format(k)
  if k % functions_per_line == functions_per_line - 1: # last
    out_msg += '\n'
  
if portion % functions_per_line != functions_per_line - 1:
  out_msg+= '\n'

out_msg += '\nDebug.Print "START"\n\n'
out_msg += 'Dim buf() As String\n'
out_msg += 'buf = Split(arr, ",")\n'

out_msg += 'Dim lenn As Long\n'
out_msg += 'lenn = ubound(buf)-lbound(buf)+1\n'
out_msg += 'debug.print (lenn)\n\n'

out_msg += 'Dim buf2() As Byte\n'
out_msg += 'ReDim buf2(lenn)\n'
out_msg += 'For i = 0 To lenn - 1\n'
out_msg += ' buf2(i) = buf(i)\n'
out_msg += 'Next i\n\n'

out_msg += 'open "vba_gen.txt" for binary access write as #1\n' 
out_msg += 'lWritePos=1\n'
out_msg += 'put #1, lWritePos, buf2\n'
out_msg += 'close #1\n'

out_msg += 'debug.print "ENDL"\n\n'

out_msg += '\nend sub'

f=open('output.txt','w')
f.write(out_msg)
f.close()
