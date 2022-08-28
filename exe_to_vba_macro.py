import sys

bytes_per_line = 30
lines_per_function = 50
bytes_per_function = bytes_per_line * lines_per_function
functions_per_line = 15

def read_file_as_bytes(fn):
  f = open(fn, 'rb')
  s = f.read()
  f.close()
  return s

def write_to_file(out_msg):
  f = open('output.txt','w')
  f.write(out_msg)
  f.close()

def construct_vba_macro(s):
  out_msg = ''
  very_first_line = 1
  first_line = 1
  # Splitting up byte strings into functions due to VBA's 64k procedure limit 
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

  # Concatenate byte strings by calling all the individual functions
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

  # Writes reconstructed file whenever macro is run
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
  return out_msg

# MAIN HERE
def main():
  if len(sys.argv) < 2:
    print('[ERROR] Insufficient arguments. Usage: exe_to_vba_macro.py <full path to input file>')
    return
  
  fn = sys.argv[1]
  try:
    binary_str = read_file_as_bytes(fn)
  except FileNotFoundError:
    print('[ERROR] File not found: {fn2}'.format(fn2=fn))
    return

  print('Input file read successfully: {fn2} ({b_size} bytes)'.format(
    fn2 = fn, 
    b_size = len(binary_str)
  ))

  # Construct VBA macro by providing binary string of file
  out_msg = construct_vba_macro(binary_str)
  
  # Write VBA macro to file
  write_to_file(out_msg)
  print('VBA macro output to: {fn2}'.format(fn2='output.txt'))

if __name__ == "__main__":
  main()
