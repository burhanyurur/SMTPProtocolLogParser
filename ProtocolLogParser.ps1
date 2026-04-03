<#
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

<#
SMTP Protocol Log Parser v1.0 - Initial release
- Parses SMTP protocol logs from Microsoft Exchange Server and generates an HTML report with detailed analysis and statistics
- Provides a GUI for selecting log files, viewing parsed data, and exporting reports
- Developed by CloudVision (https://www.cloudvision.com.tr)
#>
#Requires -Version 5.1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ================================================================
#  GLOBAL STATE
# ================================================================
$Global:Sessions = [System.Collections.Hashtable]::Synchronized(@{})
$Global:LogPath  = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) `
                              ("ProtocolLogParser_" + (Get-Date -Format "yyyyMMdd") + ".log")
$Global:ActiveParseTimer = $null
$Global:ActiveParseState = $null
$Global:LogoBase64 = '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 '

function Get-LogoBitmap {
    $bytes = [System.Convert]::FromBase64String($Global:LogoBase64)
    $ms    = [System.IO.MemoryStream]::new($bytes, 0, $bytes.Length)
    return [System.Drawing.Bitmap]::new($ms)
}

# ================================================================
#  LOGGING
# ================================================================
function WriteAppLog {
    param([string]$Level, [string]$Msg)
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"), $Level, $Msg
    try {
        $sw = [System.IO.StreamWriter]::new($Global:LogPath, $true, [System.Text.Encoding]::UTF8)
        $sw.WriteLine($line)
        $sw.Close()
    } catch {}
}

# ================================================================
#  CSV PARSER  (handles quoted fields containing commas)
# ================================================================
function ParseCsvLine {
    param([string]$Line)
    $out = [System.Collections.Generic.List[string]]::new()
    $i   = 0
    $len = $Line.Length
    while ($i -lt $len) {
        if ($Line[$i] -eq '"') {
            $i++
            $sb = [System.Text.StringBuilder]::new()
            while ($i -lt $len) {
                if ($Line[$i] -eq '"') {
                    if (($i + 1) -lt $len -and $Line[$i+1] -eq '"') { [void]$sb.Append('"'); $i += 2 }
                    else                                              { $i++; break }
                } else { [void]$sb.Append($Line[$i]); $i++ }
            }
            $out.Add($sb.ToString())
            if ($i -lt $len -and $Line[$i] -eq ',') { $i++ }
        } else {
            $start = $i
            while ($i -lt $len -and $Line[$i] -ne ',') { $i++ }
            $out.Add($Line.Substring($start, $i - $start))
            if ($i -lt $len) { $i++ }
        }
    }
    return $out
}

# ================================================================
#  SESSION PARSER
# ================================================================
function New-Session {
    param([string]$Id, [string]$Connector, [string]$Local, [string]$Remote)
    @{
        SessionId      = $Id
        ConnectorId    = $Connector
        LocalEndpoint  = $Local
        RemoteEndpoint = $Remote
        StartTime      = $null
        EndTime        = $null
        IsComplete     = $false
        SenderAddress  = ''   # first mail sender - for grid/search compat
        Recipients     = [System.Collections.Generic.List[string]]::new()  # first mail recipients
        MessageId      = ''   # first mail message-id
        Status         = 'Incomplete'
        ErrorCode      = ''
        ErrorMessage   = ''
        HasMail        = $false
        Mails          = [System.Collections.Generic.List[hashtable]]::new()
        Entries        = [System.Collections.Generic.List[hashtable]]::new()
        EhloHost       = ''
        TlsInfo        = @{ Used=$false; Protocol=''; Cipher=''; CipherBits=0; Mac=''; KeyExchange=''; CertSubject=''; CertIssuer=''; DomainCaps=''; Status='' }
    }
}

function New-MailObject {
    param([int]$StartSeq)
    @{
        SenderAddress = ''
        Recipients    = [System.Collections.Generic.List[string]]::new()
        MessageId     = ''
        StartSeq      = $StartSeq
        EndSeq        = -1
        Status        = 'Incomplete'
        ErrorCode     = ''
        ErrorMessage  = ''
    }
}

# Returns @{ OK=$true/false; Reason=''; Warning='' }
# OK=false means the file must be skipped; Warning means it loaded but something looks off.
function Test-LogFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return @{ OK=$false; Reason="File not found: $Path"; Warning='' }
    }

    $info = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($null -eq $info) {
        return @{ OK=$false; Reason="Cannot access file: $Path"; Warning='' }
    }
    if ($info.Length -eq 0) {
        return @{ OK=$false; Reason="File is empty: $($info.Name)"; Warning='' }
    }

    # Read the first ~40 lines to find the #Fields header and sample a data line
    $reader = $null
    try {
        $reader = [System.IO.StreamReader]::new($Path, [System.Text.Encoding]::UTF8)
        $fieldsLine  = ''
        $hasDataLine = $false
        $lineNum     = 0
        while ($lineNum -lt 40) {
            $line = $reader.ReadLine()
            if ($null -eq $line) { break }
            $lineNum++
            if ($line.StartsWith('#Fields:')) { $fieldsLine = $line; continue }
            if ($fieldsLine -ne '' -and -not $line.StartsWith('#') -and $line.Trim() -ne '') {
                $hasDataLine = $true; break
            }
        }

        if ($fieldsLine -eq '') {
            return @{ OK=$false; Reason="No #Fields header found in $($info.Name). This does not look like an Exchange SMTP Receive log."; Warning='' }
        }

        # Check required columns exist
        $required = @('date-time','connector-id','session-id','sequence-number','local-endpoint','remote-endpoint','event','data')
        $missing  = $required | Where-Object { $fieldsLine -notlike "*$_*" }
        if ($missing.Count -gt 0) {
            return @{ OK=$false; Reason="$($info.Name) is missing required columns: $($missing -join ', '). Wrong log type?"; Warning='' }
        }

        $warn = if (-not $hasDataLine) { "No data lines found in $($info.Name) - file may be empty or header-only." } else { '' }
        return @{ OK=$true; Reason=''; Warning=$warn }
    } catch {
        return @{ OK=$false; Reason="Cannot read $($info.Name): $($_.Exception.Message)"; Warning='' }
    } finally {
        if ($null -ne $reader) { $reader.Close() }
    }
}

function Get-FileLineCount {
    param([string]$Path)

    $reader = $null
    $lineCount = 0
    try {
        $reader = [System.IO.StreamReader]::new($Path, [System.Text.Encoding]::UTF8)
        while ($null -ne $reader.ReadLine()) {
            $lineCount++
        }
    } finally {
        if ($null -ne $reader) { $reader.Close() }
    }
    return $lineCount
}


function ProcessProtocolLogLine {
    param(
        [hashtable]$Sessions,
        [string]$RawLine,
        [ref]$HeaderFound,
        [ref]$ErrorCount
    )

    if ($RawLine.StartsWith('#')) {
        if ($RawLine -match '^#Fields:\s*(.+)$') { $HeaderFound.Value = $true }
        return
    }
    if (-not $HeaderFound.Value -or $RawLine.Trim() -eq '') { return }

    try {
        $f = @(ParseCsvLine $RawLine)
        if ($f.Count -lt 7) { $ErrorCount.Value++; return }

        $dt        = $f[0]
        $connector = $f[1]
        $sessionId = $f[2]
        $localEp   = $f[4]
        $remoteEp  = $f[5]
        $eventcode = $f[6]
        $data      = if ($f.Count -gt 7) { $f[7] } else { '' }
        $context   = if ($f.Count -gt 8) { $f[8] } else { '' }

        if (-not $Sessions.ContainsKey($sessionId)) {
            $Sessions[$sessionId] = New-Session $sessionId $connector $localEp $remoteEp
        }
        $s = $Sessions[$sessionId]
        if ($s.ConnectorId -eq '' -and $connector -ne '') { $s.ConnectorId = $connector }

        $s.Entries.Add(@{ DateTime=$dt; SequenceNumber=$f[3]; Event=$eventcode; Data=$data; Context=$context })

        switch ($eventcode) {
            '+' { $s.StartTime = $dt; $s.RemoteEndpoint = $remoteEp; $s.LocalEndpoint = $localEp }
            '-' { $s.EndTime = $dt; $s.IsComplete = $true
                  if ($s.Status -eq 'Incomplete') { $s.Status = 'Complete' } }
            '<' {
                if ($data -match '(?i)^EHLO\s+(\S+)') {
                    if ($s.EhloHost -eq '') { $s.EhloHost = $Matches[1] }
                }
                elseif ($data -match '(?i)^MAIL FROM:\s*<?([^>]+)>?') {
                    $senderaddress = $Matches[1].Trim()
                    $mail   = New-MailObject ([int]$f[3])
                    $mail.SenderAddress = $senderaddress
                    $s.Mails.Add($mail)
                    $s.HasMail = $true
                    if ($s.SenderAddress -eq '') { $s.SenderAddress = $senderaddress }
                }
                elseif ($data -match '(?i)^RCPT TO:\s*<?([^>]+)>?') {
                    $rcpt = $Matches[1].Trim()
                    if ($s.Mails.Count -gt 0) { $s.Mails[$s.Mails.Count - 1].Recipients.Add($rcpt) }
                    if ($s.Mails.Count -le 1)  { $s.Recipients.Add($rcpt) }
                }
            }
            '>' {
                if ($data -match '^([45]\d{2})\s+(.*)$') {
                    $code = $Matches[1]; $msg = $Matches[2].TrimEnd()
                    if ($s.Mails.Count -gt 0) {
                        $lm = $s.Mails[$s.Mails.Count - 1]
                        if ($lm.ErrorCode -eq '') { $lm.ErrorCode = $code; $lm.ErrorMessage = $msg; $lm.Status = 'Error' }
                    }
                    $s.ErrorCode = $code; $s.ErrorMessage = $msg; $s.Status = 'Error'
                }
            }
            '*' {
                if ($context -eq 'receiving message' -and $data -match '^([^;]+);') {
                    $msgId = $Matches[1]
                    if ($s.Mails.Count -gt 0 -and $s.Mails[$s.Mails.Count - 1].MessageId -eq '') {
                        $s.Mails[$s.Mails.Count - 1].MessageId = $msgId
                    }
                    if ($s.MessageId -eq '') { $s.MessageId = $msgId }
                }
                # TLS: certificate subject (context field)
                if ($context -eq 'Certificate subject') {
                    $s.TlsInfo.Used = $true; $s.TlsInfo.CertSubject = $data
                }
                elseif ($context -eq 'Certificate issuer name') {
                    $s.TlsInfo.CertIssuer = $data
                }
                # TLS: negotiation success line (in context column)
                elseif ($context -match 'TLS protocol\s+(\S+)\s+negotiation succeeded.*?algorithm\s+(\w+)\s+with strength\s+(\d+).*?MAC hash algorithm\s+(\w+).*?key exchange algorithm\s+(\w+)') {
                    $s.TlsInfo.Used = $true
                    $s.TlsInfo.Protocol    = $Matches[1]
                    $s.TlsInfo.Cipher      = $Matches[2]
                    $s.TlsInfo.CipherBits  = [int]$Matches[3]
                    $s.TlsInfo.Mac         = $Matches[4]
                    $s.TlsInfo.KeyExchange = $Matches[5]
                }
                # TLS domain capabilities (in context column)
                elseif ($context -match "TlsDomainCapabilities='([^']+)'.*?Status='([^']+)'") {
                    $s.TlsInfo.DomainCaps = $Matches[1]; $s.TlsInfo.Status = $Matches[2]
                }
            }
        }
    } catch {
        $ErrorCount.Value++
    }
}

function FinalizeParsedSessions {
    param([hashtable]$Sessions)

    foreach ($s in $Sessions.Values) {
        # Set EndSeq and status for each mail object
        for ($i = 0; $i -lt $s.Mails.Count; $i++) {
            $m = $s.Mails[$i]
            if ($i + 1 -lt $s.Mails.Count) {
                $m.EndSeq = $s.Mails[$i + 1].StartSeq - 1
            } elseif ($s.Entries.Count -gt 0) {
                $m.EndSeq = [int]$s.Entries[$s.Entries.Count - 1].SequenceNumber
            }
            if ($m.Status -eq 'Incomplete' -and $m.MessageId -ne '') { $m.Status = 'OK' }
        }
        if ($s.Status -eq 'Complete' -and $s.SenderAddress -ne '' -and $s.ErrorCode -eq '') {
            $s.Status = 'OK'
        }
    }
}

function New-ParseOperationState {
    param([string[]]$FilePaths)

    WriteAppLog 'INFO' "Counting lines for $($FilePaths.Count) file(s)"
    $lineTotals = @{}
    $totalLines = 0
    foreach ($path in $FilePaths) {
        $lineCount = 0
        try { $lineCount = Get-FileLineCount $path } catch {}
        $lineTotals[$path] = $lineCount
        $totalLines += $lineCount
    }

    @{
        Sessions         = [System.Collections.Hashtable]::Synchronized(@{})
        FilePaths        = $FilePaths
        FileTotal        = $FilePaths.Count
        FileIndex        = -1
        CurrentPath      = ''
        CurrentFileName  = ''
        CurrentFileLines = 0
        CurrentFileTotal = 0
        CurrentErrors    = 0
        HeaderFound      = $false
        Reader           = $null
        LineTotals       = $lineTotals
        TotalLines       = $totalLines
        CompletedLines   = 0
        ProcessedLines   = 0
        RemainingLines   = $totalLines
        Phase            = 'Preparing'
        IsComplete       = $false
    }
}

function Open-NextParseFile {
    param([hashtable]$State)

    if ($null -ne $State.Reader) {
        $State.Reader.Close()
        $State.Reader = $null
    }

    $State.FileIndex++
    if ($State.FileIndex -ge $State.FileTotal) {
        $State.IsComplete = $true
        $State.Phase = 'Completed'
        return $false
    }

    $path = $State.FilePaths[$State.FileIndex]
    $State.CurrentPath = $path
    $State.CurrentFileName = Split-Path $path -Leaf
    $State.CurrentFileLines = 0
    $State.CurrentErrors = 0
    $State.HeaderFound = $false
    $State.CurrentFileTotal = if ($State.LineTotals.ContainsKey($path)) { [int]$State.LineTotals[$path] } else { 0 }
    $State.Reader = [System.IO.StreamReader]::new($path, [System.Text.Encoding]::UTF8)
    $State.Phase = 'Parsing'
    WriteAppLog 'INFO' "Parsing: $path"
    return $true
}

function Invoke-ParseBatch {
    param(
        [hashtable]$State,
        [int]$BatchSize = 1000
    )

    if ($State.IsComplete) { return }

    $processedThisTick = 0
    while ($processedThisTick -lt $BatchSize) {
        if ($null -eq $State.Reader) {
            if (-not (Open-NextParseFile $State)) {
                FinalizeParsedSessions $State.Sessions
                $State.ProcessedLines = $State.CompletedLines
                $State.RemainingLines = 0
                return
            }
        }

        $raw = $State.Reader.ReadLine()
        if ($null -eq $raw) {
            WriteAppLog 'INFO' "Done: $($State.CurrentPath)  Lines:$($State.CurrentFileLines)  Errors:$($State.CurrentErrors)"
            $State.CompletedLines += $State.CurrentFileLines
            $State.ProcessedLines = $State.CompletedLines
            $State.RemainingLines = [Math]::Max(0, $State.TotalLines - $State.ProcessedLines)
            $State.Phase = 'Completed'
            $State.Reader.Close()
            $State.Reader = $null
            continue
        }

        $State.CurrentFileLines++
        $processedThisTick++
        # [ref] only works on local variables, not hashtable slots - copy, pass, write back
        $headerVal = $State.HeaderFound
        $errorVal  = $State.CurrentErrors
        $headerRef = [ref]$headerVal
        $errorRef  = [ref]$errorVal
        ProcessProtocolLogLine -Sessions $State.Sessions -RawLine $raw -HeaderFound $headerRef -ErrorCount $errorRef
        $State.HeaderFound   = $headerRef.Value
        $State.CurrentErrors = $errorRef.Value

        $State.ProcessedLines = $State.CompletedLines + $State.CurrentFileLines
        $State.RemainingLines = [Math]::Max(0, $State.TotalLines - $State.ProcessedLines)
        $State.Phase = 'Parsing'
    }
}

function ParseLogFiles {
    param([string[]]$FilePaths, [System.ComponentModel.BackgroundWorker]$Worker = $null)

    $sessions       = [System.Collections.Hashtable]::Synchronized(@{})
    $fileIndex      = 0
    $fileTotal      = $FilePaths.Count
    $lineTotals     = @{}
    $totalLines     = 0
    $completedLines = 0

    foreach ($path in $FilePaths) {
        $lineCount = 0
        try { $lineCount = Get-FileLineCount $path } catch {}
        $lineTotals[$path] = $lineCount
        $totalLines += $lineCount
    }

    if ($null -ne $Worker) {
        $Worker.ReportProgress(0, @{
            FileIndex      = 0
            FileTotal      = $fileTotal
            LineCount      = 0
            FileName       = ''
            FileLineTotal  = 0
            ProcessedLines = 0
            CompletedLines = 0
            RemainingLines = $totalLines
            Phase          = 'Preparing'
        })
    }

    foreach ($path in $FilePaths) {
        $fileIndex++
        WriteAppLog 'INFO' "Parsing: $path"
        $lines     = 0
        $errors    = 0
        $fileName  = Split-Path $path -Leaf
        $fileLineTotal = if ($lineTotals.ContainsKey($path)) { [int]$lineTotals[$path] } else { 0 }
        try {
            $reader = [System.IO.StreamReader]::new($path, [System.Text.Encoding]::UTF8)
            $headerFound = $false
            while ($null -ne ($raw = $reader.ReadLine())) {
                $lines++
                if ($null -ne $Worker -and ($lines -eq 1 -or $lines % 500 -eq 0 -or $lines -eq $fileLineTotal)) {
                    $processedLines = $completedLines + $lines
                    $remainingLines = [Math]::Max(0, $totalLines - $processedLines)
                    $pct = if ($totalLines -gt 0) { [int](($processedLines * 100) / $totalLines) } else { 0 }
                    $Worker.ReportProgress($pct, @{
                        FileIndex      = $fileIndex
                        FileTotal      = $fileTotal
                        LineCount      = $lines
                        FileName       = $fileName
                        FileLineTotal  = $fileLineTotal
                        ProcessedLines = $processedLines
                        CompletedLines = $completedLines
                        RemainingLines = $remainingLines
                        Phase          = 'Parsing'
                    })
                }
                if ($raw.StartsWith('#')) {
                    if ($raw -match '^#Fields:\s*(.+)$') { $headerFound = $true }
                    continue
                }
                if (-not $headerFound -or $raw.Trim() -eq '') { continue }
                try {
                    # @() ensures a mutable PS array regardless of what ParseCsvLine returns
                    $f = @(ParseCsvLine $raw)
                    if ($f.Count -lt 7) { $errors++; continue }

                    $dt        = $f[0]
                    $connector = $f[1]
                    $sessionId = $f[2]
                    # $seqNum  = $f[3]  (stored per-entry below)
                    $localEp   = $f[4]
                    $remoteEp  = $f[5]
                    $eventcode     = $f[6]
                    $data      = if ($f.Count -gt 7) { $f[7] } else { '' }
                    $context   = if ($f.Count -gt 8) { $f[8] } else { '' }

                    if (-not $sessions.ContainsKey($sessionId)) {
                        $sessions[$sessionId] = New-Session $sessionId $connector $localEp $remoteEp
                    }
                    $s = $sessions[$sessionId]
                    if ($s.ConnectorId -eq '' -and $connector -ne '') { $s.ConnectorId = $connector }

                    $s.Entries.Add(@{ DateTime=$dt; SequenceNumber=$f[3]; Event=$eventcode; Data=$data; Context=$context })

                    switch ($eventcode) {
                        '+' { $s.StartTime = $dt; $s.RemoteEndpoint = $remoteEp; $s.LocalEndpoint = $localEp }
                        '-' { $s.EndTime = $dt; $s.IsComplete = $true
                              if ($s.Status -eq 'Incomplete') { $s.Status = 'Complete' } }
                        '<' {
                            if      ($data -match '(?i)^MAIL FROM:\s*<?([^>]+)>?')  {
                                $s.SenderAddress = $Matches[1].Trim(); $s.HasMail = $true
                            }
                            elseif  ($data -match '(?i)^RCPT TO:\s*<?([^>]+)>?')    {
                                $s.Recipients.Add($Matches[1].Trim())
                            }
                        }
                        '>' {
                            if ($data -match '^([45]\d{2})\s+(.*)$') {
                                $s.ErrorCode    = $Matches[1]
                                $s.ErrorMessage = $Matches[2].TrimEnd()
                                $s.Status       = 'Error'
                            }
                        }
                        '*' {
                            if ($context -eq 'receiving message' -and $s.MessageId -eq '' -and $data -match '^([^;]+);') {
                                $s.MessageId = $Matches[1]
                            }
                        }
                    }
                } catch { $errors++ }
            }
            $reader.Close()
        } catch { WriteAppLog 'ERROR' "File read failed: $path - $_" }

        WriteAppLog 'INFO' "Done: $path  Lines:$lines  Errors:$errors"
        $completedLines += $lines
        if ($null -ne $Worker) {
            $remainingLines = [Math]::Max(0, $totalLines - $completedLines)
            $pct = if ($totalLines -gt 0) { [int](($completedLines * 100) / $totalLines) } else { 100 }
            $Worker.ReportProgress($pct, @{
                FileIndex      = $fileIndex
                FileTotal      = $fileTotal
                LineCount      = $lines
                FileName       = $fileName
                FileLineTotal  = $fileLineTotal
                ProcessedLines = $completedLines
                CompletedLines = $completedLines
                RemainingLines = $remainingLines
                Phase          = 'Completed'
            })
        }
    }

    # Finalize status for sessions that completed with a sender but no error
    foreach ($s in $sessions.Values) {
        if ($s.Status -eq 'Complete' -and $s.SenderAddress -ne '' -and $s.ErrorCode -eq '') {
            $s.Status = 'OK'
        }
    }
    return $sessions
}

# ================================================================
#  STATISTICS
# ================================================================
function Get-Statistics {
    param($Sessions)
    $senders  = @{}; $receivers = @{}; $errors = @{}; $byHour = @{}
    $status   = @{ OK=0; Error=0; Incomplete=0; Complete=0 }
    $ehloHosts   = @{}
    $tlsProtocols = @{}
    $tlsCiphers  = @{}
    $tlsCount    = 0
    $noTlsCount  = 0

    foreach ($s in $Sessions.Values) {
        $k = $s.Status
        if ($status.ContainsKey($k)) { $status[$k]++ }

        if ($s.SenderAddress -ne '') {
            if (-not $senders[$s.SenderAddress]) { $senders[$s.SenderAddress] = 0 }
            $senders[$s.SenderAddress]++
        }
        foreach ($r in $s.Recipients) {
            if (-not $receivers[$r]) { $receivers[$r] = 0 }
            $receivers[$r]++
        }
        if ($s.ErrorCode -ne '') {
            $ek = "$($s.ErrorCode) - $($s.ErrorMessage.Substring(0,[Math]::Min(50,$s.ErrorMessage.Length)))"
            if (-not $errors[$ek]) { $errors[$ek] = 0 }
            $errors[$ek]++
        }
        if ($null -ne $s.StartTime -and $s.StartTime.Length -ge 13) {
            $hr = $s.StartTime.Substring(11,2) + ':00'
            if (-not $byHour[$hr]) { $byHour[$hr] = 0 }
            $byHour[$hr]++
        }

        # EHLO hosts
        if ($s.EhloHost -ne '') {
            if (-not $ehloHosts[$s.EhloHost]) { $ehloHosts[$s.EhloHost] = 0 }
            $ehloHosts[$s.EhloHost]++
        }

        # TLS
        if ($s.TlsInfo.Used) {
            $tlsCount++
            if ($s.TlsInfo.Protocol -ne '') {
                if (-not $tlsProtocols[$s.TlsInfo.Protocol]) { $tlsProtocols[$s.TlsInfo.Protocol] = 0 }
                $tlsProtocols[$s.TlsInfo.Protocol]++
            }
            if ($s.TlsInfo.Cipher -ne '') {
                $ck = if ($s.TlsInfo.CipherBits -gt 0) { "$($s.TlsInfo.Cipher) ($($s.TlsInfo.CipherBits)-bit)" } else { $s.TlsInfo.Cipher }
                if (-not $tlsCiphers[$ck]) { $tlsCiphers[$ck] = 0 }
                $tlsCiphers[$ck]++
            }
        } else {
            $noTlsCount++
        }
    }
    return @{
        TopSenders    = ($senders.GetEnumerator()    | Sort-Object Value -Desc | Select-Object -First 10)
        TopReceivers  = ($receivers.GetEnumerator()   | Sort-Object Value -Desc | Select-Object -First 10)
        TopErrors     = ($errors.GetEnumerator()      | Sort-Object Value -Desc | Select-Object -First 10)
        ByHour        = ($byHour.GetEnumerator()      | Sort-Object Key)
        StatusCounts  = $status
        TotalSessions = $Sessions.Count
        TotalMails    = ($Sessions.Values | Where-Object { $_.HasMail } | Measure-Object).Count
        TopEhloHosts  = ($ehloHosts.GetEnumerator()   | Sort-Object Value -Desc | Select-Object -First 15)
        TlsCount      = $tlsCount
        NoTlsCount    = $noTlsCount
        TlsProtocols  = $tlsProtocols
        TopTlsCiphers = ($tlsCiphers.GetEnumerator()  | Sort-Object Value -Desc | Select-Object -First 10)
    }
}

# ================================================================
#  GDI+ CHARTS
# ================================================================
function New-BarChart {
    param(
        [string]$Title,
        [object[]]$Data,
        [string]$KeyProp   = 'Key',
        [string]$ValProp   = 'Value',
        [int]$Width        = 650,
        [int]$Height       = 300,
        [System.Drawing.Color]$BarColor = [System.Drawing.Color]::SteelBlue
    )
    $bmp = [System.Drawing.Bitmap]::new($Width, $Height)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::White)

    $padL=65; $padR=20; $padT=34; $padB=70
    $cW = $Width - $padL - $padR
    $cH = $Height - $padT - $padB

    $tfont = [System.Drawing.Font]::new('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
    $g.DrawString($Title, $tfont, [System.Drawing.Brushes]::DarkSlateBlue, [float]($padL), [float]6)

    if ($null -eq $Data -or $Data.Count -eq 0) {
        $g.DrawString('No data available', [System.Drawing.Font]::new('Segoe UI',9),
            [System.Drawing.Brushes]::Gray, [float]($padL + 20), [float]($padT + $cH/2))
        $g.Dispose(); $tfont.Dispose(); return $bmp
    }

    $maxVal = ($Data | ForEach-Object { [int]$_.$ValProp } | Measure-Object -Maximum).Maximum
    if ($maxVal -lt 1) { $maxVal = 1 }

    $cnt   = $Data.Count
    $slotW = $cW / $cnt
    $barW  = [Math]::Max(4, $slotW - 6)

    $barBrush  = [System.Drawing.SolidBrush]::new($BarColor)
    $axisPen   = [System.Drawing.Pen]::new([System.Drawing.Color]::Gray, 1)
    $lbFont    = [System.Drawing.Font]::new('Segoe UI', 7.5)
    $valFont   = [System.Drawing.Font]::new('Segoe UI', 7)

    # axes
    $g.DrawLine($axisPen, $padL, $padT, $padL, $padT + $cH)
    $g.DrawLine($axisPen, $padL, $padT + $cH, $padL + $cW, $padT + $cH)

    # y-axis tick labels
    $axisValFont = [System.Drawing.Font]::new('Segoe UI', 7)
    for ($tick = 0; $tick -le 4; $tick++) {
        $tv  = [int]($maxVal * $tick / 4)
        $ty  = [float]($padT + $cH - ($cH * $tick / 4))
        $g.DrawString($tv.ToString(), $axisValFont, [System.Drawing.Brushes]::Gray, [float]2, $ty - 7)
        $g.DrawLine([System.Drawing.Pens]::LightGray, [float]$padL, $ty, [float]($padL + $cW), $ty)
    }

    for ($i = 0; $i -lt $cnt; $i++) {
        $item  = $Data[$i]
        $val   = [int]$item.$ValProp
        $bh    = [int](([double]$val / $maxVal) * $cH)
        $bx    = [int]($padL + $i * $slotW + ($slotW - $barW) / 2)
        $by    = [int]($padT + $cH - $bh)

        $g.FillRectangle($barBrush, $bx, $by, [int]$barW, $bh)

        # value on top
        $vStr = $val.ToString()
        $vSz  = $g.MeasureString($vStr, $valFont)
        $g.DrawString($vStr, $valFont, [System.Drawing.Brushes]::DimGray,
            [float]($bx + $barW/2 - $vSz.Width/2), [float]([Math]::Max($padT - 1, $by - $vSz.Height)))

        # rotated label on x-axis
        $lbl = $item.$KeyProp.ToString()
        if ($lbl.Length -gt 15) { $lbl = $lbl.Substring(0,14) + '..' }
        $lx  = [float]($bx + $barW/2)
        $ly  = [float]($padT + $cH + 4)
        $state = $g.Save()
        $g.TranslateTransform($lx, $ly)
        $g.RotateTransform(35)
        $g.DrawString($lbl, $lbFont, [System.Drawing.Brushes]::DimGray, [float]0, [float]0)
        $g.Restore($state)
    }

    $barBrush.Dispose(); $axisPen.Dispose(); $lbFont.Dispose()
    $valFont.Dispose(); $tfont.Dispose(); $axisValFont.Dispose()
    $g.Dispose()
    return $bmp
}

function New-PieChart {
    param(
        [string]$Title,
        [hashtable]$Data,
        [int]$Width  = 460,
        [int]$Height = 300
    )
    $bmp = [System.Drawing.Bitmap]::new($Width, $Height)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::White)

    $tf = [System.Drawing.Font]::new('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
    $g.DrawString($Title, $tf, [System.Drawing.Brushes]::DarkSlateBlue, [float]10, [float]6)

    $total = 0; foreach ($v in $Data.Values) { $total += $v }

    if ($total -eq 0) {
        $g.DrawString('No data', [System.Drawing.Font]::new('Segoe UI',9),
            [System.Drawing.Brushes]::Gray, [float]20, [float]($Height/2))
        $g.Dispose(); $tf.Dispose(); return $bmp
    }

    $colors = @(
        [System.Drawing.Color]::SteelBlue,  [System.Drawing.Color]::Tomato,
        [System.Drawing.Color]::SeaGreen,   [System.Drawing.Color]::Orange,
        [System.Drawing.Color]::MediumPurple,[System.Drawing.Color]::Goldenrod
    )
    $lf = [System.Drawing.Font]::new('Segoe UI', 8)
    $px = 20; $py = 36; $pw = 220; $ph = 220
    $angle = -90.0
    $ci = 0
    foreach ($kv in $Data.GetEnumerator()) {
        if ($kv.Value -le 0) { $ci++; continue }
        $sweep = [float](($kv.Value / $total) * 360.0)
        $br    = [System.Drawing.SolidBrush]::new($colors[$ci % $colors.Count])
        $g.FillPie($br, [float]$px,[float]$py,[float]$pw,[float]$ph, $angle, $sweep)
        $g.DrawPie([System.Drawing.Pens]::White,[float]$px,[float]$py,[float]$pw,[float]$ph,$angle,$sweep)
        $ly = 40 + $ci * 22
        $g.FillRectangle($br, [float]($px+$pw+16), [float]$ly, [float]14, [float]14)
        $g.DrawString("$($kv.Key): $($kv.Value)", $lf, [System.Drawing.Brushes]::Black,
            [float]($px+$pw+34), [float]$ly)
        $br.Dispose()
        $angle += $sweep; $ci++
    }
    $tf.Dispose(); $lf.Dispose(); $g.Dispose()
    return $bmp
}

function BitmapToBase64 {
    param([System.Drawing.Bitmap]$Bmp)
    $ms = [System.IO.MemoryStream]::new()
    $Bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    [System.Convert]::ToBase64String($ms.ToArray())
}

# ================================================================
#  HTML EXPORT
# ================================================================
function HtmlEncode {
    param([string]$s)
    $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;')
}

function Export-HtmlReport {
    param([string]$OutputPath, $Sessions, $Stats)

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<!DOCTYPE html><html><head><meta charset="UTF-8"><title>SMTP Protocol Log Parser v1.0 Report</title><style>
body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin:0;background:#f5f5f5;color:#222}
h1{background:#1a3a5c;color:#fff;padding:12px 22px;margin:0;font-size:20px}
h2{color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:3px;margin-top:28px}
.container{padding:20px 28px}
table{border-collapse:collapse;width:100%;margin-bottom:20px}
th{background:#1a3a5c;color:#fff;padding:7px 10px;text-align:left;font-weight:600}
td{padding:5px 10px;border-bottom:1px solid #ddd;vertical-align:top}
tr:nth-child(even) td{background:#f0f4fa}
tr:hover td{background:#dce8f8}
.sb{display:inline-block;background:#fff;border:1px solid #ccc;border-radius:6px;
    padding:12px 20px;margin:6px;text-align:center;min-width:110px;box-shadow:0 1px 3px #0001}
.sn{font-size:28px;font-weight:bold;color:#1a3a5c}
.sl{font-size:11px;color:#666;margin-top:2px}
.charts{display:flex;flex-wrap:wrap;gap:14px;margin-bottom:20px}
img.chart{border:1px solid #ccc;border-radius:4px;background:#fff}
.ok{color:#2a7a2a;font-weight:600}.er{color:#c0392b;font-weight:600}.ic{color:#d35400;font-weight:600}
.hdr{display:flex;align-items:center;justify-content:space-between;background:#1a3a5c;padding:8px 22px}
.hdr h1{color:#fff;font-size:20px;margin:0}
.hdr img{height:44px}
.ftr{text-align:center;padding:18px;color:#888;font-size:11px;border-top:1px solid #ddd;margin-top:20px}
.ftr a{color:#1a6a9a;text-decoration:none}
</style></head><body>')
    $logoB64Html = $Global:LogoBase64
    [void]$sb.Append("<div class='hdr'><h1>SMTP Protocol Log - Analysis Report</h1><a href='https://www.cloudvision.com.tr'><img src='data:image/png;base64,$logoB64Html' />></a>)</div><div class='container'>")

    # summary
    [void]$sb.Append('<h2>Summary</h2><div>')
    $boxes = @(
        @{L='Total Sessions'; V=$Stats.TotalSessions},
        @{L='Emails (MAIL FROM)'; V=$Stats.TotalMails},
        @{L='OK'; V=$Stats.StatusCounts.OK},
        @{L='Errors'; V=$Stats.StatusCounts.Error},
        @{L='Incomplete'; V=$Stats.StatusCounts.Incomplete}
    )
    foreach ($b in $boxes) {
        [void]$sb.Append("<div class='sb'><div class='sn'>$($b.V)</div><div class='sl'>$($b.L)</div></div>")
    }
    [void]$sb.Append('</div>')

    # charts
    [void]$sb.Append('<h2>Statistics</h2><div class="charts">')
    $c1 = New-BarChart 'Top Senders'    $Stats.TopSenders    -Width 660 -Height 310
    $c2 = New-BarChart 'Top Recipients' $Stats.TopReceivers   -Width 660 -Height 310 -BarColor ([System.Drawing.Color]::SeaGreen)
    $c3 = New-BarChart 'Top Errors'     $Stats.TopErrors      -Width 660 -Height 310 -BarColor ([System.Drawing.Color]::Tomato)
    $hourArr = $Stats.ByHour | ForEach-Object { [PSCustomObject]@{Key=$_.Key;Value=$_.Value} }
    $c4 = New-BarChart 'Sessions by Hour' $hourArr            -Width 760 -Height 310 -BarColor ([System.Drawing.Color]::DarkSlateBlue)
    $sd  = @{}; foreach ($kv in $Stats.StatusCounts.GetEnumerator()) { if ($kv.Value -gt 0) { $sd[$kv.Key] = $kv.Value } }
    $c5 = New-PieChart 'Session Status Distribution' $sd      -Width 480 -Height 310
    $c6 = New-BarChart 'Top EHLO Hosts (by session count)' $Stats.TopEhloHosts -Width 760 -Height 310 -BarColor ([System.Drawing.Color]::CadetBlue)
    $tlsPieData = @{}
    if ($Stats.TlsCount -gt 0)   { $tlsPieData['TLS'] = $Stats.TlsCount }
    if ($Stats.NoTlsCount -gt 0) { $tlsPieData['No TLS'] = $Stats.NoTlsCount }
    $c7 = New-PieChart 'TLS Usage' $tlsPieData -Width 480 -Height 310
    $tlsProtoArr = $Stats.TlsProtocols.GetEnumerator() | Sort-Object Value -Desc | ForEach-Object { [PSCustomObject]@{Key=$_.Key;Value=$_.Value} }
    $c8 = New-BarChart 'TLS Protocol Versions' $tlsProtoArr  -Width 560 -Height 310 -BarColor ([System.Drawing.Color]::DarkCyan)
    $c9 = New-BarChart 'Top TLS Cipher Suites' $Stats.TopTlsCiphers -Width 760 -Height 310 -BarColor ([System.Drawing.Color]::SlateBlue)
    foreach ($c in @($c1,$c2,$c3,$c4,$c5,$c6,$c7,$c8,$c9)) {
        [void]$sb.Append("<img class='chart' src='data:image/png;base64,$(BitmapToBase64 $c)' />")
        $c.Dispose()
    }
    [void]$sb.Append('</div>')

    # top senders table
    [void]$sb.Append('<h2>Top Senders</h2><table><tr><th>Sender</th><th>Count</th></tr>')
    foreach ($kv in $Stats.TopSenders) {
        [void]$sb.Append("<tr><td>$(HtmlEncode $kv.Key)</td><td>$($kv.Value)</td></tr>")
    }
    [void]$sb.Append('</table>')

    # top receivers table
    [void]$sb.Append('<h2>Top Recipients</h2><table><tr><th>Recipient</th><th>Count</th></tr>')
    foreach ($kv in $Stats.TopReceivers) {
        [void]$sb.Append("<tr><td>$(HtmlEncode $kv.Key)</td><td>$($kv.Value)</td></tr>")
    }
    [void]$sb.Append('</table>')

    # errors table
    [void]$sb.Append('<h2>Error Summary</h2><table><tr><th>Error</th><th>Count</th></tr>')
    foreach ($kv in $Stats.TopErrors) {
        [void]$sb.Append("<tr><td class='er'>$(HtmlEncode $kv.Key)</td><td>$($kv.Value)</td></tr>")
    }
    [void]$sb.Append('</table>')

    # EHLO hosts table
    [void]$sb.Append('<h2>Top EHLO Hosts</h2><table><tr><th>EHLO Hostname</th><th>Sessions</th></tr>')
    foreach ($kv in $Stats.TopEhloHosts) {
        [void]$sb.Append("<tr><td>$(HtmlEncode $kv.Key)</td><td>$($kv.Value)</td></tr>")
    }
    [void]$sb.Append('</table>')

    # TLS summary table
    $tlsTotal = $Stats.TlsCount + $Stats.NoTlsCount
    $tlsPct   = if ($tlsTotal -gt 0) { [int](($Stats.TlsCount / $tlsTotal) * 100) } else { 0 }
    [void]$sb.Append('<h2>TLS Summary</h2><table><tr><th>Metric</th><th>Value</th></tr>')
    [void]$sb.Append("<tr><td>Sessions with TLS</td><td>$($Stats.TlsCount) ($tlsPct%)</td></tr>")
    [void]$sb.Append("<tr><td>Sessions without TLS</td><td>$($Stats.NoTlsCount)</td></tr>")
    [void]$sb.Append('</table>')

    if ($Stats.TlsProtocols.Count -gt 0) {
        [void]$sb.Append('<table><tr><th>TLS Protocol</th><th>Sessions</th></tr>')
        foreach ($kv in ($Stats.TlsProtocols.GetEnumerator() | Sort-Object Value -Desc)) {
            [void]$sb.Append("<tr><td>$(HtmlEncode $kv.Key)</td><td>$($kv.Value)</td></tr>")
        }
        [void]$sb.Append('</table>')
    }

    if ($Stats.TopTlsCiphers.Count -gt 0) {
        [void]$sb.Append('<table><tr><th>Cipher Suite</th><th>Sessions</th></tr>')
        foreach ($kv in $Stats.TopTlsCiphers) {
            [void]$sb.Append("<tr><td>$(HtmlEncode $kv.Key)</td><td>$($kv.Value)</td></tr>")
        }
        [void]$sb.Append('</table>')
    }

    # sessions table
    [void]$sb.Append('<h2>All Sessions</h2><table><tr><th>Session ID</th><th>Connector</th><th>Remote IP</th><th>Start</th><th>EHLO Host</th><th>TLS</th><th>Sender</th><th>Recipients</th><th>Status</th><th>Error</th></tr>')
    foreach ($s in ($Sessions.Values | Sort-Object { $_.StartTime })) {
        $cls     = switch ($s.Status) { 'OK' {'ok'} 'Error' {'er'} default {'ic'} }
        $ip      = if ($s.RemoteEndpoint -match '^(.+):\d+$') { $Matches[1] } else { $s.RemoteEndpoint }
        $tlsCell = if ($s.TlsInfo.Used) {
            $proto = if ($s.TlsInfo.Protocol -ne '') { $s.TlsInfo.Protocol } else { 'Yes' }
            "<span style='color:#2a7a2a;font-weight:600'>$proto</span>"
        } else { "<span style='color:#888'>No</span>" }
        $ehloCell = HtmlEncode ($(if ($s.EhloHost -ne '') { $s.EhloHost } else { '' }))
        [void]$sb.Append("<tr><td>$($s.SessionId)</td><td>$(HtmlEncode $s.ConnectorId)</td><td>$ip</td><td>$($s.StartTime)</td><td>$ehloCell</td><td>$tlsCell</td><td>$(HtmlEncode $s.SenderAddress)</td><td>$(HtmlEncode ($s.Recipients -join '; '))</td><td class='$cls'>$($s.Status)</td><td class='er'>$(HtmlEncode $s.ErrorCode)</td></tr>")
    }
    [void]$sb.Append('</table></div></body></html>')

    [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
}

# ================================================================
#  UI HELPERS
# ================================================================
function Get-RemoteIP {
    param([string]$Ep)
    if ($Ep -match '^(.+):\d+$') { return $Matches[1] }
    return $Ep
}

function Get-ConnectorGroups {
    param($Sessions)
    $g = @{}
    foreach ($s in $Sessions.Values) {
        $k = if ($s.ConnectorId -ne '') { $s.ConnectorId } else { '(No Connector)' }
        if (-not $g.ContainsKey($k)) { $g[$k] = [System.Collections.Generic.List[object]]::new() }
        $g[$k].Add($s)
    }
    return $g
}

function PopulateTreeView {
    param($TV, $Sessions)
    $TV.BeginUpdate(); $TV.Nodes.Clear()
    $groups = Get-ConnectorGroups $Sessions
    foreach ($conn in ($groups.Keys | Sort-Object)) {
        $cn = [System.Windows.Forms.TreeNode]::new($conn)
        $cn.Tag = @{ Type='Connector'; ConnectorId=$conn }
        $cn.ForeColor = [System.Drawing.Color]::DarkBlue
        $cn.NodeFont  = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        foreach ($s in ($groups[$conn] | Sort-Object { $_.StartTime })) {
            $ip      = Get-RemoteIP $s.RemoteEndpoint
            $mailCnt = if ($s.Mails.Count -gt 0) { "  [$($s.Mails.Count) mail(s)]" } else { '' }
            $sn      = [System.Windows.Forms.TreeNode]::new("$($s.SessionId)  [$ip]  $($s.StartTime)  [$($s.Status)]$mailCnt")
            $sn.Tag  = @{ Type='Session'; Session=$s }
            $sn.ForeColor = switch ($s.Status) {
                'Error'      { [System.Drawing.Color]::Firebrick }
                'Incomplete' { [System.Drawing.Color]::DarkOrange }
                default      { [System.Drawing.Color]::DarkGreen }
            }
            foreach ($mail in $s.Mails) {
                $rcpts = if ($mail.Recipients.Count -gt 0) { $mail.Recipients -join ', ' } else { '(no recipients)' }
                $mn = [System.Windows.Forms.TreeNode]::new("$($mail.SenderAddress) -> $rcpts")
                $mn.Tag = @{ Type='Mail'; Session=$s; Mail=$mail }
                $mn.ForeColor = switch ($mail.Status) {
                    'Error'      { [System.Drawing.Color]::Firebrick }
                    'Incomplete' { [System.Drawing.Color]::DarkOrange }
                    default      { [System.Drawing.Color]::Navy }
                }
                [void]$sn.Nodes.Add($mn)
            }
            [void]$cn.Nodes.Add($sn)
        }
        [void]$TV.Nodes.Add($cn)
    }
    $TV.EndUpdate()
}

function FilterTreeView {
    param($TV, $Sessions, [string]$Filter)
    if ($Filter -eq '') { PopulateTreeView $TV $Sessions; return }
    $fl = $Filter.ToLower()
    $TV.BeginUpdate(); $TV.Nodes.Clear()
    $groups = Get-ConnectorGroups $Sessions
    foreach ($conn in ($groups.Keys | Sort-Object)) {
        $cn = $null
        foreach ($s in $groups[$conn]) {
            $ip      = Get-RemoteIP $s.RemoteEndpoint
            $allAddr = ($s.Mails | ForEach-Object { $_.SenderAddress; $_.Recipients }) -join ' '
            $txt     = "$($s.SessionId) $ip $allAddr".ToLower()
            if ($txt -notlike "*$fl*") { continue }
            if ($null -eq $cn) {
                $cn = [System.Windows.Forms.TreeNode]::new($conn)
                $cn.Tag = @{ Type='Connector'; ConnectorId=$conn }
                $cn.ForeColor = [System.Drawing.Color]::DarkBlue
                $cn.NodeFont  = [System.Drawing.Font]::new('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
            }
            $sn = [System.Windows.Forms.TreeNode]::new("$($s.SessionId)  [$ip]  $($s.StartTime)  [$($s.Status)]")
            $sn.Tag = @{ Type='Session'; Session=$s }
            $sn.ForeColor = switch ($s.Status) {
                'Error' { [System.Drawing.Color]::Firebrick }
                'Incomplete' { [System.Drawing.Color]::DarkOrange }
                default { [System.Drawing.Color]::DarkGreen }
            }
            [void]$cn.Nodes.Add($sn)
        }
        if ($null -ne $cn) { [void]$TV.Nodes.Add($cn) }
    }
    $TV.EndUpdate()
}

function PopulateEhloTree {
    param($TV, $Sessions)
    $TV.BeginUpdate(); $TV.Nodes.Clear()
    # group by EhloHost
    $groups = @{}
    foreach ($s in $Sessions.Values) {
        $ehlo = if ($s.EhloHost -ne '') { $s.EhloHost } else { '(no EHLO)' }
        if (-not $groups.ContainsKey($ehlo)) { $groups[$ehlo] = [System.Collections.Generic.List[hashtable]]::new() }
        $groups[$ehlo].Add($s)
    }
    foreach ($ehlo in ($groups.Keys | Sort-Object)) {
        $hn = [System.Windows.Forms.TreeNode]::new("$ehlo  [$($groups[$ehlo].Count) session(s)]")
        $hn.Tag = @{ Type='EhloHost'; Host=$ehlo }
        $hn.ForeColor = [System.Drawing.Color]::DarkBlue
        $hn.NodeFont  = [System.Drawing.Font]::new('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
        foreach ($s in ($groups[$ehlo] | Sort-Object { $_.StartTime })) {
            $ip      = Get-RemoteIP $s.RemoteEndpoint
            $mailCnt = if ($s.Mails.Count -gt 0) { "  [$($s.Mails.Count) mail(s)]" } else { '' }
            $sn = [System.Windows.Forms.TreeNode]::new("$($s.SessionId)  [$ip]  $($s.StartTime)  [$($s.Status)]$mailCnt")
            $sn.Tag = @{ Type='Session'; Session=$s }
            $sn.ForeColor = switch ($s.Status) {
                'Error'      { [System.Drawing.Color]::Firebrick }
                'Incomplete' { [System.Drawing.Color]::DarkOrange }
                default      { [System.Drawing.Color]::DarkGreen }
            }
            foreach ($mail in $s.Mails) {
                $rcpts = if ($mail.Recipients.Count -gt 0) { $mail.Recipients -join ', ' } else { '(no recipients)' }
                $mn = [System.Windows.Forms.TreeNode]::new("$($mail.SenderAddress) -> $rcpts")
                $mn.Tag = @{ Type='Mail'; Session=$s; Mail=$mail }
                $mn.ForeColor = switch ($mail.Status) {
                    'Error'      { [System.Drawing.Color]::Firebrick }
                    'Incomplete' { [System.Drawing.Color]::DarkOrange }
                    default      { [System.Drawing.Color]::Navy }
                }
                [void]$sn.Nodes.Add($mn)
            }
            [void]$hn.Nodes.Add($sn)
        }
        [void]$TV.Nodes.Add($hn)
    }
    $TV.EndUpdate()
}

function PopulateTlsTree {
    param($TV, $Sessions)
    $TV.BeginUpdate(); $TV.Nodes.Clear()
    foreach ($s in ($Sessions.Values | Sort-Object { $_.StartTime })) {
        $ip  = Get-RemoteIP $s.RemoteEndpoint
        $tls = $s.TlsInfo
        $tlsLabel = if ($tls.Used) {
            if ($tls.Protocol -ne '') { "TLS: $($tls.Protocol)  $($tls.Cipher)/$($tls.CipherBits)-bit" }
            else                       { 'TLS: negotiated' }
        } else { 'No TLS' }
        $sn = [System.Windows.Forms.TreeNode]::new("$($s.SessionId)  [$ip]  $($s.StartTime)  [$tlsLabel]")
        $sn.Tag = @{ Type='Session'; Session=$s }
        $sn.ForeColor = if ($tls.Used) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DimGray }

        if ($tls.Used) {
            $fields = [ordered]@{
                'Protocol'      = $tls.Protocol
                'Cipher'        = "$($tls.Cipher) ($($tls.CipherBits) bit)"
                'MAC'           = $tls.Mac
                'Key Exchange'  = $tls.KeyExchange
                'Cert Subject'  = $tls.CertSubject
                'Cert Issuer'   = $tls.CertIssuer
                'Domain Caps'   = $tls.DomainCaps
                'Status'        = $tls.Status
            }
            foreach ($kv in $fields.GetEnumerator()) {
                if ($kv.Value -ne '' -and $kv.Value -ne ' (0 bit)') {
                    $fn = [System.Windows.Forms.TreeNode]::new("$($kv.Key): $($kv.Value)")
                    $fn.Tag = @{ Type='TlsDetail' }
                    $fn.ForeColor = [System.Drawing.Color]::DarkSlateGray
                    [void]$sn.Nodes.Add($fn)
                }
            }
        }
        foreach ($mail in $s.Mails) {
            $rcpts = if ($mail.Recipients.Count -gt 0) { $mail.Recipients -join ', ' } else { '(no recipients)' }
            $mn = [System.Windows.Forms.TreeNode]::new("$($mail.SenderAddress) -> $rcpts")
            $mn.Tag = @{ Type='Mail'; Session=$s; Mail=$mail }
            $mn.ForeColor = switch ($mail.Status) {
                'Error'      { [System.Drawing.Color]::Firebrick }
                'Incomplete' { [System.Drawing.Color]::DarkOrange }
                default      { [System.Drawing.Color]::Navy }
            }
            [void]$sn.Nodes.Add($mn)
        }
        [void]$TV.Nodes.Add($sn)
    }
    $TV.EndUpdate()
}

function Set-GridCols {
    param($Global:grid, [string[]]$Cols, [int[]]$Widths)
    $Global:grid.Columns.Clear()
    for ($i = 0; $i -lt $Cols.Count; $i++) {
        $c = [System.Windows.Forms.DataGridViewTextBoxColumn]::new()
        $c.Name = $Cols[$i]; $c.HeaderText = $Cols[$i]
        if ($i -lt $Widths.Count) { $c.Width = $Widths[$i] }
        else { $c.Width = 120 }
        $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        [void]$Global:grid.Columns.Add($c)
    }
}

function PopulateGridConnector {
    param($Global:grid, $Sessions, [string]$ConnId)
    $Global:grid.Rows.Clear()
    Set-GridCols $Global:grid `
        @('Session-ID','Remote IP','Start Time','End Time','Sender','Recipients','Status','Error','Error Message') `
        @(140,110,145,145,160,200,75,65,200)
    $rows = $Sessions.Values | Where-Object { $null -eq $ConnId -or $_.ConnectorId -eq $ConnId }
    foreach ($s in ($rows | Sort-Object { $_.StartTime })) {
        $ri = $Global:grid.Rows.Add($s.SessionId, (Get-RemoteIP $s.RemoteEndpoint), $s.StartTime,
              $s.EndTime, $s.SenderAddress, ($s.Recipients -join '; '),
              $s.Status, $s.ErrorCode, $s.ErrorMessage)
        $Global:grid.Rows[$ri].Tag = $s
        $Global:grid.Rows[$ri].DefaultCellStyle.BackColor = switch ($s.Status) {
            'Error'      { [System.Drawing.Color]::FromArgb(255,235,235) }
            'Incomplete' { [System.Drawing.Color]::FromArgb(255,250,225) }
            default      { [System.Drawing.Color]::White }
        }
    }
}

function PopulateGridSession {
    param($Global:grid, $Session)
    $Global:grid.Rows.Clear()
    Set-GridCols $Global:grid @('Seq#','Time','Event','Data','Context') @(50,145,50,440,200)
    foreach ($e in $Session.Entries) {
        $ri = $Global:grid.Rows.Add($e.SequenceNumber, $e.DateTime, $e.Event, $e.Data, $e.Context)
        $Global:grid.Rows[$ri].DefaultCellStyle.BackColor = switch ($e.Event) {
            '+' { [System.Drawing.Color]::FromArgb(225,255,225) }
            '-' { [System.Drawing.Color]::FromArgb(255,225,225) }
            '>' { [System.Drawing.Color]::FromArgb(235,245,255) }
            '<' { [System.Drawing.Color]::FromArgb(255,250,235) }
            '*' { [System.Drawing.Color]::FromArgb(248,248,248) }
            default { [System.Drawing.Color]::White }
        }
    }
}

function PopulateGridMail {
    param($Global:grid, $Session, $Mail = $null)
    $Global:grid.Rows.Clear()
    Set-GridCols $Global:grid @('Seq#','Time','Event','Data','Context') @(50,145,50,440,200)
    foreach ($e in $Session.Entries) {
        if ($null -ne $Mail) {
            $seq = [int]$e.SequenceNumber
            if ($seq -lt $Mail.StartSeq) { continue }
            if ($Mail.EndSeq -ge 0 -and $seq -gt $Mail.EndSeq) { continue }
        }
        $ri = $Global:grid.Rows.Add($e.SequenceNumber, $e.DateTime, $e.Event, $e.Data, $e.Context)
        $Global:grid.Rows[$ri].DefaultCellStyle.BackColor = switch ($e.Event) {
            '+' { [System.Drawing.Color]::FromArgb(225,255,225) }
            '-' { [System.Drawing.Color]::FromArgb(255,225,225) }
            '>' { [System.Drawing.Color]::FromArgb(235,245,255) }
            '<' { [System.Drawing.Color]::FromArgb(255,250,235) }
            '*' { [System.Drawing.Color]::FromArgb(248,248,248) }
            default { [System.Drawing.Color]::White }
        }
    }
}

function Update-DetailPanel {
    param($RTB, $Entry)
    $RTB.Clear()
    if ($null -eq $Entry) { return }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("Time     : $($Entry.DateTime)")
    $lines.Add("Seq#     : $($Entry.SequenceNumber)")
    $lines.Add("Event    : $($Entry.Event)")
    $lines.Add('')
    switch ($Entry.Event) {
        '+' { $lines.Add('[Connection Opened]') }
        '-' { $lines.Add("[Connection Closed]"); if ($Entry.Context -ne '') { $lines.Add("Reason   : $($Entry.Context)") } }
        '>' {
            $lines.Add('[Server -> Client Response]')
            if ($Entry.Data -match '^(\d{3})[-\s](.*)$') {
                $lines.Add("Code     : $($Matches[1])"); $lines.Add("Message  : $($Matches[2].TrimEnd())")
            } else { $lines.Add("Data     : $($Entry.Data)") }
        }
        '<' {
            $lines.Add('[Client -> Server Command]')
            if ($Entry.Data -match '^(\S+)\s*(.*)$') {
                $lines.Add("Command  : $($Matches[1].ToUpper())")
                if ($Matches[2] -ne '') { $lines.Add("Argument : $($Matches[2].TrimEnd())") }
            } else { $lines.Add("Data     : $($Entry.Data)") }
        }
        '*' {
            $lines.Add('[Internal Event]')
            $lines.Add("Context  : $($Entry.Context)")
            if ($Entry.Data -match 'SMTP') {
                $lines.Add(''); $lines.Add('Permission Flags:')
                foreach ($flag in ($Entry.Data -split '\s+' | Where-Object { $_ -ne '' })) {
                    $lines.Add("  - $flag")
                }
            } else { $lines.Add("Data     : $($Entry.Data)") }
        }
    }
    $RTB.Text = $lines -join "`r`n"
}

function Build-StatsTab {
    param($Panel, $Stats)
    $Panel.Controls.Clear()

    $flow = [System.Windows.Forms.FlowLayoutPanel]::new()
    $flow.Dock = [System.Windows.Forms.DockStyle]::Top; $flow.Height = 76
    $flow.Padding = [System.Windows.Forms.Padding]::new(8,6,0,0)
    $sumItems = @(
        @{L='Total Sessions';    V=$Stats.TotalSessions},
        @{L='Emails (MAIL FROM)';V=$Stats.TotalMails},
        @{L='OK';                V=$Stats.StatusCounts.OK},
        @{L='Errors';            V=$Stats.StatusCounts.Error},
        @{L='Incomplete';        V=$Stats.StatusCounts.Incomplete}
    )
    foreach ($item in $sumItems) {
        $gb  = [System.Windows.Forms.GroupBox]::new(); $gb.Text=$item.L; $gb.Width=120; $gb.Height=60
        $lbl = [System.Windows.Forms.Label]::new()
        $lbl.Text = $item.V.ToString()
        $lbl.Font = [System.Drawing.Font]::new('Segoe UI',16,[System.Drawing.FontStyle]::Bold)
        $lbl.ForeColor = [System.Drawing.Color]::DarkBlue
        $lbl.Dock = [System.Windows.Forms.DockStyle]::Fill
        $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $gb.Controls.Add($lbl); [void]$flow.Controls.Add($gb)
    }
    [void]$Panel.Controls.Add($flow)

    $scroll = [System.Windows.Forms.Panel]::new()
    $scroll.Dock = [System.Windows.Forms.DockStyle]::Fill
    $scroll.AutoScroll = $true

    $y = 10
    $chartDefs = @(
        @{ T='Top Senders';      D=$Stats.TopSenders;   C=[System.Drawing.Color]::SteelBlue;     W=680; H=310 },
        @{ T='Top Recipients';   D=$Stats.TopReceivers; C=[System.Drawing.Color]::SeaGreen;      W=680; H=310 },
        @{ T='Top Errors';       D=$Stats.TopErrors;    C=[System.Drawing.Color]::Tomato;        W=680; H=310 },
        @{ T='Sessions by Hour'; D=($Stats.ByHour | ForEach-Object { [PSCustomObject]@{Key=$_.Key;Value=$_.Value} });
                                                        C=[System.Drawing.Color]::DarkSlateBlue; W=800; H=310 }
    )
    foreach ($cd in $chartDefs) {
        $bmp = New-BarChart $cd.T $cd.D -Width $cd.W -Height $cd.H -BarColor $cd.C
        $pb  = [System.Windows.Forms.PictureBox]::new()
        $pb.Image = $bmp; $pb.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::AutoSize
        $pb.Location = [System.Drawing.Point]::new(10, $y)
        [void]$scroll.Controls.Add($pb); $y += $cd.H + 14
    }
    $sd = @{}
    foreach ($kv in $Stats.StatusCounts.GetEnumerator()) { if ($kv.Value -gt 0) { $sd[$kv.Key] = $kv.Value } }
    $pie = New-PieChart 'Session Status Distribution' $sd -Width 480 -Height 310
    $pp  = [System.Windows.Forms.PictureBox]::new()
    $pp.Image = $pie; $pp.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::AutoSize
    $pp.Location = [System.Drawing.Point]::new(10, $y)
    [void]$scroll.Controls.Add($pp)

    [void]$Panel.Controls.Add($scroll)
}

function Build-ErrorsTab {
    param($Panel, $Sessions)
    $Panel.Controls.Clear()

    $toolbar = [System.Windows.Forms.FlowLayoutPanel]::new()
    $toolbar.Dock = [System.Windows.Forms.DockStyle]::Top; $toolbar.Height = 36
    $toolbar.Padding = [System.Windows.Forms.Padding]::new(8,5,0,0)

    $lbl = [System.Windows.Forms.Label]::new(); $lbl.Text = 'Group by:'; $lbl.AutoSize = $true
    $lbl.Margin = [System.Windows.Forms.Padding]::new(0,4,6,0)
    [void]$toolbar.Controls.Add($lbl)

    $combo = [System.Windows.Forms.ComboBox]::new()
    $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$combo.Items.AddRange(@('Error Code','Sender Address','Sender IP'))
    $combo.SelectedIndex = 0; $combo.Width = 150
    [void]$toolbar.Controls.Add($combo)
    $egrid = [System.Windows.Forms.DataGridView]::new()
    $egrid.Dock = [System.Windows.Forms.DockStyle]::Fill; $egrid.ReadOnly = $true
    $egrid.AllowUserToAddRows = $false; $egrid.RowHeadersVisible = $false
    $egrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $egrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $egrid.RowTemplate.Height = 22
    [void]$Panel.Controls.Add($egrid)
    [void]$Panel.Controls.Add($toolbar)

    $fillGrid = {
        param([string]$by, $egr, $sess)
        $egr.Rows.Clear(); $egr.Columns.Clear()
        $errSess = $sess.Values | Where-Object { $_.ErrorCode -ne '' -or $_.Status -eq 'Error' }
        switch ($by) {
            'Error Code' {
                Set-GridCols $egr @('Error Code','Count','Sample Message','Sender Addresses','Sender IPs') @(80,55,260,220,160)
                $g = @{}
                foreach ($s in $errSess) {
                    $k = if ($s.ErrorCode -ne '') { $s.ErrorCode } else { '???' }
                    if (-not $g[$k]) { $g[$k] = @{Cnt=0;Msg=$s.ErrorMessage;SA=@{};SI=@{}} }
                    $g[$k].Cnt++
                    if ($s.SenderAddress -ne '') { $g[$k].SA[$s.SenderAddress]=1 }
                    $ip = Get-RemoteIP $s.RemoteEndpoint; if ($ip -ne '') { $g[$k].SI[$ip]=1 }
                }
                foreach ($kv in ($g.GetEnumerator() | Sort-Object { $_.Value.Cnt } -Desc)) {
                    [void]$egr.Rows.Add($kv.Key, $kv.Value.Cnt, $kv.Value.Msg,
                        ($kv.Value.SA.Keys -join '; '), ($kv.Value.SI.Keys -join '; '))
                }
            }
            'Sender Address' {
                Set-GridCols $egr @('Sender Address','Error Count','Error Codes') @(240,80,300)
                $g = @{}
                foreach ($s in $errSess) {
                    $k = if ($s.SenderAddress -ne '') { $s.SenderAddress } else { '(unknown)' }
                    if (-not $g[$k]) { $g[$k] = @{Cnt=0;Codes=@{}} }
                    $g[$k].Cnt++; if ($s.ErrorCode -ne '') { $g[$k].Codes[$s.ErrorCode]=1 }
                }
                foreach ($kv in ($g.GetEnumerator() | Sort-Object { $_.Value.Cnt } -Desc)) {
                    [void]$egr.Rows.Add($kv.Key, $kv.Value.Cnt, ($kv.Value.Codes.Keys -join ', '))
                }
            }
            'Sender IP' {
                Set-GridCols $egr @('Sender IP','Error Count','Error Codes') @(160,80,300)
                $g = @{}
                foreach ($s in $errSess) {
                    $k = Get-RemoteIP $s.RemoteEndpoint; if ($k -eq '') { $k = '(unknown)' }
                    if (-not $g[$k]) { $g[$k] = @{Cnt=0;Codes=@{}} }
                    $g[$k].Cnt++; if ($s.ErrorCode -ne '') { $g[$k].Codes[$s.ErrorCode]=1 }
                }
                foreach ($kv in ($g.GetEnumerator() | Sort-Object { $_.Value.Cnt } -Desc)) {
                    [void]$egr.Rows.Add($kv.Key, $kv.Value.Cnt, ($kv.Value.Codes.Keys -join ', '))
                }
            }
        }
        # color error rows red-tint
        foreach ($row in $egr.Rows) {
            if ($row.Cells.Count -gt 0 -and $row.Cells[0].Value -match '^[45]\d{2}') {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick
            }
        }
    }

    & $fillGrid $combo.SelectedItem.ToString() $egrid $Sessions

    $combo.Add_SelectedIndexChanged({
        & $fillGrid $combo.SelectedItem.ToString() $egrid $Sessions
    }.GetNewClosure())
}

# ================================================================
#  MAIN FORM
# ================================================================
function Build-MainForm {
    $form = [System.Windows.Forms.Form]::new()
    $form.Text = "SMTP Protocol Log Parser v1.0 - Exchange Server"
    $form.Size = [System.Drawing.Size]::new(1300, 820)
    $form.MinimumSize = [System.Drawing.Size]::new(1000, 620)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.Font = [System.Drawing.Font]::new('Segoe UI', 9)

    # ---- MenuStrip ----
    $menu    = [System.Windows.Forms.MenuStrip]::new()
    $miFile  = [System.Windows.Forms.ToolStripMenuItem]::new('File')
    $miOpen  = [System.Windows.Forms.ToolStripMenuItem]::new('Open Log File(s)...')
    $miOpen.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::O
    $miExport = [System.Windows.Forms.ToolStripMenuItem]::new('Export HTML Report...')
    $miExport.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::E
    $miExport.Enabled = $false
    $miViewLog = [System.Windows.Forms.ToolStripMenuItem]::new('View Application Log')
    $miExit   = [System.Windows.Forms.ToolStripMenuItem]::new('Exit')
    [void]$miFile.DropDownItems.AddRange(@($miOpen, $miExport,
        [System.Windows.Forms.ToolStripSeparator]::new(), $miViewLog,
        [System.Windows.Forms.ToolStripSeparator]::new(), $miExit))
    $miHelp  = [System.Windows.Forms.ToolStripMenuItem]::new('Help')
    $miAbout = [System.Windows.Forms.ToolStripMenuItem]::new('About')
    [void]$miHelp.DropDownItems.Add($miAbout)
    [void]$menu.Items.AddRange(@($miFile, $miHelp))
    $form.MainMenuStrip = $menu

    # ---- ToolStrip ----
    $toolbar  = [System.Windows.Forms.ToolStrip]::new()
    $tbOpen   = [System.Windows.Forms.ToolStripButton]::new('Open Files')
    $tbExport = [System.Windows.Forms.ToolStripButton]::new('Export HTML'); $tbExport.Enabled = $false
    $tbLog    = [System.Windows.Forms.ToolStripButton]::new('View Log')
    $tbProg   = [System.Windows.Forms.ToolStripProgressBar]::new(); $tbProg.Width=0
    $tbLabel  = [System.Windows.Forms.ToolStripLabel]::new('Ready')
    $logoBmp  = Get-LogoBitmap
    $logoH    = $toolbar.Height - 4; if ($logoH -lt 20) { $logoH = 28 }
    $logoW    = [int]($logoBmp.Width * $logoH / $logoBmp.Height)
    $logoScaled = [System.Drawing.Bitmap]::new($logoBmp, [System.Drawing.Size]::new($logoW, $logoH))
    $logoBmp.Dispose()
    $tbLogo   = [System.Windows.Forms.ToolStripLabel]::new()
    $tbLogo.Image        = $logoScaled
    $tbLogo.ImageScaling = [System.Windows.Forms.ToolStripItemImageScaling]::None
    $tbLogo.Alignment    = [System.Windows.Forms.ToolStripItemAlignment]::Right
    $tbLogo.Padding      = [System.Windows.Forms.Padding]::new(6, 0, 6, 0)
    [void]$toolbar.Items.AddRange(@($tbOpen,$tbExport,
        [System.Windows.Forms.ToolStripSeparator]::new(),$tbLog,
        [System.Windows.Forms.ToolStripSeparator]::new(),$tbProg,$tbLabel,$tbLogo))

    # ---- StatusStrip ----
    $statusBar  = [System.Windows.Forms.StatusStrip]::new()
    $sbMsg      = [System.Windows.Forms.ToolStripStatusLabel]::new('Ready')
    $sbMsg.Spring = $true; $sbMsg.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $sbSessions = [System.Windows.Forms.ToolStripStatusLabel]::new('Sessions: 0')
    $sbFiles    = [System.Windows.Forms.ToolStripStatusLabel]::new('Files: 0')
    $sbTime     = [System.Windows.Forms.ToolStripStatusLabel]::new('')
    $sbLink     = [System.Windows.Forms.ToolStripStatusLabel]::new('www.cloudvision.com.tr')
    $sbLink.IsLink    = $true
    $sbLink.LinkColor = [System.Drawing.Color]::SteelBlue
    $sbLink.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
    $sbLink.Add_Click({ Start-Process 'https://www.cloudvision.com.tr' })
    [void]$statusBar.Items.AddRange(@($sbMsg,$sbSessions,$sbFiles,$sbTime,$sbLink))

    # ---- Layout ----
    $mainSplit = [System.Windows.Forms.SplitContainer]::new()
    $mainSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainSplit.SplitterDistance = 330
    $mainSplit.Panel1MinSize = 200

    # Left: TabControl with Sessions / EHLO / TLS trees
    $leftTabs = [System.Windows.Forms.TabControl]::new()
    $leftTabs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $leftTabs.Font = [System.Drawing.Font]::new('Segoe UI', 8.5)

    # helper to make a tree panel (search box + label + treeview inside a tab page)
    $makeTreeTab = {
        param([string]$TabTitle, [string]$SearchHint)
        $tp = [System.Windows.Forms.TabPage]::new($TabTitle)
        $tp.Padding = [System.Windows.Forms.Padding]::new(0)

        $lbl = [System.Windows.Forms.Label]::new()
        $lbl.Text = $SearchHint; $lbl.Dock = [System.Windows.Forms.DockStyle]::Top
        $lbl.Height = 18; $lbl.Font = [System.Drawing.Font]::new('Segoe UI',7.5)
        $lbl.Padding = [System.Windows.Forms.Padding]::new(4,2,0,0)
        $lbl.ForeColor = [System.Drawing.Color]::DimGray

        $sb = [System.Windows.Forms.TextBox]::new()
        $sb.Dock = [System.Windows.Forms.DockStyle]::Top; $sb.Height = 24
        $sb.Font = [System.Drawing.Font]::new('Segoe UI',9)

        $tv = [System.Windows.Forms.TreeView]::new()
        $tv.Dock = [System.Windows.Forms.DockStyle]::Fill
        $tv.HideSelection = $false
        $tv.Font = [System.Drawing.Font]::new('Consolas', 8.5)
        $tv.ItemHeight = 20

        [void]$tp.Controls.Add($tv)
        [void]$tp.Controls.Add($sb)
        [void]$tp.Controls.Add($lbl)
        return @{ Page=$tp; Tree=$tv; Search=$sb }
    }

    $tabSessions = & $makeTreeTab 'Sessions' 'Filter sessions:'
    $tabEhlo     = & $makeTreeTab 'EHLO'     'Filter by EHLO host:'
    $tabTls      = & $makeTreeTab 'TLS'      'Filter TLS sessions:'

    $Global:treeView  = $tabSessions.Tree
    $tvSearch         = $tabSessions.Search
    $ehloTree         = $tabEhlo.Tree
    $ehloSearch       = $tabEhlo.Search
    $tlsTree          = $tabTls.Tree
    $tlsSearch        = $tabTls.Search

    [void]$leftTabs.TabPages.AddRange(@($tabSessions.Page, $tabEhlo.Page, $tabTls.Page))
    [void]$mainSplit.Panel1.Controls.Add($leftTabs)

    # Right: TabControl
    $tabs       = [System.Windows.Forms.TabControl]::new()
    $tabs.Dock  = [System.Windows.Forms.DockStyle]::Fill

    # Tab 1: Protocol View
    $tabProto = [System.Windows.Forms.TabPage]::new('Protocol View')
    $pSplit   = [System.Windows.Forms.SplitContainer]::new()
    $pSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pSplit.Orientation = [System.Windows.Forms.Orientation]::Horizontal

    $pSplit.Panel2MinSize = 80

    $Global:grid= [System.Windows.Forms.DataGridView]::new()
    $Global:grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $Global:grid.ReadOnly = $true; $Global:grid.AllowUserToAddRows = $false
    $Global:grid.RowHeadersVisible = $false; $Global:grid.MultiSelect = $false
    $Global:grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $Global:grid.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $Global:grid.RowTemplate.Height = 22
    $Global:grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $Global:grid.ColumnHeadersHeight = 26
    $Global:grid.EnableHeadersVisualStyles = $false
    $Global:grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(26,58,92)
    $Global:grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $Global:grid.ColumnHeadersDefaultCellStyle.Font = [System.Drawing.Font]::new('Segoe UI',9,[System.Drawing.FontStyle]::Bold)

    $detailRtb = [System.Windows.Forms.RichTextBox]::new()
    $detailRtb.Dock = [System.Windows.Forms.DockStyle]::Fill; $detailRtb.ReadOnly = $true
    $detailRtb.Font = [System.Drawing.Font]::new('Consolas', 9)
    $detailRtb.BackColor = [System.Drawing.Color]::FromArgb(248,250,255)

    [void]$pSplit.Panel1.Controls.Add($Global:grid)
    [void]$pSplit.Panel2.Controls.Add($detailRtb)
    [void]$tabProto.Controls.Add($pSplit)

    # Tab 2: Statistics
    $tabStats   = [System.Windows.Forms.TabPage]::new('Statistics')
    $statsPanel = [System.Windows.Forms.Panel]::new(); $statsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    [void]$tabStats.Controls.Add($statsPanel)

    # Tab 3: Errors
    $tabErrors   = [System.Windows.Forms.TabPage]::new('Errors')
    $errorsPanel = [System.Windows.Forms.Panel]::new(); $errorsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    [void]$tabErrors.Controls.Add($errorsPanel)

    # Tab 4: Search
    $tabSearch = [System.Windows.Forms.TabPage]::new('Search')

    $sTop = [System.Windows.Forms.Panel]::new()
    $sTop.Dock = [System.Windows.Forms.DockStyle]::Top; $sTop.Height = 100; $sTop.Padding = [System.Windows.Forms.Padding]::new(10,8,10,4)

    # Row 1
    $mkLbl = { param($t,$x,$y,$w) $l=[System.Windows.Forms.Label]::new(); $l.Text=$t; $l.Location=[System.Drawing.Point]::new($x,$y); $l.Width=$w; $l.TextAlign=[System.Drawing.ContentAlignment]::MiddleRight; [void]$sTop.Controls.Add($l); return $l }
    $mkTxt = { param($x,$y,$w)    $t=[System.Windows.Forms.TextBox]::new(); $t.Location=[System.Drawing.Point]::new($x,$y); $t.Width=$w; [void]$sTop.Controls.Add($t); return $t }

    [void](& $mkLbl 'Sender IP:'      0  10  100); $txtSenderIP   = & $mkTxt 104  8  180
    [void](& $mkLbl 'Sender Address:' 294 10 120); $txtSenderAddr = & $mkTxt 418  8  200
    [void](& $mkLbl 'Recipient:'      0  42  100); $txtRecipient  = & $mkTxt 104 40  180
    [void](& $mkLbl 'Session ID:'     294 42 120); $txtSessionID  = & $mkTxt 418 40  200

    $btnSearch = [System.Windows.Forms.Button]::new(); $btnSearch.Text = 'Search'
    $btnSearch.Location = [System.Drawing.Point]::new(104, 70); $btnSearch.Width = 90
    [void]$sTop.Controls.Add($btnSearch)
    $btnClear  = [System.Windows.Forms.Button]::new(); $btnClear.Text = 'Clear'
    $btnClear.Location = [System.Drawing.Point]::new(200, 70); $btnClear.Width = 80
    [void]$sTop.Controls.Add($btnClear)

    $searchGrid = [System.Windows.Forms.DataGridView]::new()
    $searchGrid.Dock = [System.Windows.Forms.DockStyle]::Fill; $searchGrid.ReadOnly = $true
    $searchGrid.AllowUserToAddRows = $false; $searchGrid.RowHeadersVisible = $false
    $searchGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $searchGrid.MultiSelect = $false; $searchGrid.RowTemplate.Height = 22
    $searchGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(26,58,92)
    $searchGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $searchGrid.EnableHeadersVisualStyles = $false

    [void]$tabSearch.Controls.Add($searchGrid)
    [void]$tabSearch.Controls.Add($sTop)

    [void]$tabs.TabPages.AddRange(@($tabProto, $tabStats, $tabErrors, $tabSearch))
    [void]$mainSplit.Panel2.Controls.Add($tabs)

    # Add controls to form (order matters for docking)
    [void]$form.Controls.Add($mainSplit)
    [void]$form.Controls.Add($statusBar)
    [void]$form.Controls.Add($toolbar)
    [void]$form.Controls.Add($menu)

    # ================================================================
    #  EVENT HANDLERS
    # ================================================================

    # Set treeview panel width on load: 25% of form width, capped at 100 chars in tree font
    $form.Add_Load({
        $charWidth = [System.Windows.Forms.TextRenderer]::MeasureText('X', $Global:treeView.Font).Width
        $maxWidth  = $charWidth * 100
        $preferred = [int]($form.ClientSize.Width * 0.25)
        $mainSplit.SplitterDistance = [Math]::Min($preferred, $maxWidth)
        $pSplit.SplitterDistance = [int]($form.ClientSize.Height * 0.65)
    }.GetNewClosure())

    # ---- Sessions tree ----
    $tvSearch.Add_TextChanged({
        if ($Global:Sessions.Count -gt 0) {
            FilterTreeView $Global:treeView $Global:Sessions $tvSearch.Text
        }
    }.GetNewClosure())

    $Global:treeView.Add_AfterSelect({
        $node = $Global:UI_treeView.SelectedNode
        if ($null -eq $node -or $null -eq $node.Tag) { return }
        if ($null -ne $Global:UI_tabs -and $Global:UI_tabs.SelectedTab -ne $Global:UI_tabProto) { $Global:UI_tabs.SelectedTab = $Global:UI_tabProto }
        $tag = $node.Tag
        switch ($tag.Type) {
            'Connector' { PopulateGridConnector $Global:UI_grid $Global:Sessions $tag.ConnectorId; if ($null -ne $Global:UI_detailRtb) { $Global:UI_detailRtb.Clear() } }
            'Session'   { PopulateGridSession   $Global:UI_grid $tag.Session;                      if ($null -ne $Global:UI_detailRtb) { $Global:UI_detailRtb.Clear() } }
            'Mail'      { PopulateGridMail       $Global:UI_grid $tag.Session $tag.Mail;            if ($null -ne $Global:UI_detailRtb) { $Global:UI_detailRtb.Clear() } }
        }
    })

    # ---- EHLO tree ----
    $ehloSearch.Add_TextChanged({
        if ($Global:Sessions.Count -eq 0) { return }
        $fl = $ehloSearch.Text.ToLower()
        if ($fl -eq '') { PopulateEhloTree $ehloTree $Global:Sessions; return }
        $ehloTree.BeginUpdate(); $ehloTree.Nodes.Clear()
        $groups = @{}
        foreach ($s in $Global:Sessions.Values) {
            $e = if ($s.EhloHost -ne '') { $s.EhloHost } else { '(no EHLO)' }
            if (-not $groups.ContainsKey($e)) { $groups[$e] = [System.Collections.Generic.List[hashtable]]::new() }
            $groups[$e].Add($s)
        }
        foreach ($ehlo in ($groups.Keys | Sort-Object)) {
            if ($ehlo.ToLower() -notlike "*$fl*" -and
                ($groups[$ehlo] | Where-Object { (Get-RemoteIP $_.RemoteEndpoint).ToLower() -like "*$fl*" -or $_.SessionId.ToLower() -like "*$fl*" }).Count -eq 0) { continue }
            $hn = [System.Windows.Forms.TreeNode]::new("$ehlo  [$($groups[$ehlo].Count) session(s)]")
            $hn.Tag = @{ Type='EhloHost'; Host=$ehlo }
            $hn.ForeColor = [System.Drawing.Color]::DarkBlue
            $hn.NodeFont  = [System.Drawing.Font]::new('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
            foreach ($s in ($groups[$ehlo] | Sort-Object { $_.StartTime })) {
                $ip = Get-RemoteIP $s.RemoteEndpoint
                $mailCnt = if ($s.Mails.Count -gt 0) { "  [$($s.Mails.Count) mail(s)]" } else { '' }
                $sn = [System.Windows.Forms.TreeNode]::new("$($s.SessionId)  [$ip]  $($s.StartTime)  [$($s.Status)]$mailCnt")
                $sn.Tag = @{ Type='Session'; Session=$s }
                $sn.ForeColor = switch ($s.Status) { 'Error' { [System.Drawing.Color]::Firebrick } 'Incomplete' { [System.Drawing.Color]::DarkOrange } default { [System.Drawing.Color]::DarkGreen } }
                foreach ($mail in $s.Mails) {
                    $rcpts = if ($mail.Recipients.Count -gt 0) { $mail.Recipients -join ', ' } else { '(no recipients)' }
                    $mn = [System.Windows.Forms.TreeNode]::new("$($mail.SenderAddress) -> $rcpts")
                    $mn.Tag = @{ Type='Mail'; Session=$s; Mail=$mail }
                    $mn.ForeColor = switch ($mail.Status) { 'Error' { [System.Drawing.Color]::Firebrick } 'Incomplete' { [System.Drawing.Color]::DarkOrange } default { [System.Drawing.Color]::Navy } }
                    [void]$sn.Nodes.Add($mn)
                }
                [void]$hn.Nodes.Add($sn)
            }
            [void]$ehloTree.Nodes.Add($hn)
        }
        $ehloTree.EndUpdate()
    }.GetNewClosure())

    $ehloTree.Add_AfterSelect({
        $node = $Global:UI_ehloTree.SelectedNode
        if ($null -eq $node -or $null -eq $node.Tag) { return }
        if ($null -ne $Global:UI_tabs -and $Global:UI_tabs.SelectedTab -ne $Global:UI_tabProto) { $Global:UI_tabs.SelectedTab = $Global:UI_tabProto }
        $tag = $node.Tag
        switch ($tag.Type) {
            'Session' { PopulateGridSession $Global:UI_grid $tag.Session; if ($null -ne $Global:UI_detailRtb) { $Global:UI_detailRtb.Clear() } }
            'Mail'    { PopulateGridMail    $Global:UI_grid $tag.Session $tag.Mail; if ($null -ne $Global:UI_detailRtb) { $Global:UI_detailRtb.Clear() } }
        }
    })

    # ---- TLS tree ----
    $tlsSearch.Add_TextChanged({
        if ($Global:Sessions.Count -eq 0) { return }
        $fl = $tlsSearch.Text.ToLower()
        if ($fl -eq '') { PopulateTlsTree $tlsTree $Global:Sessions; return }
        $tlsTree.BeginUpdate(); $tlsTree.Nodes.Clear()
        foreach ($s in ($Global:Sessions.Values | Sort-Object { $_.StartTime })) {
            $ip = Get-RemoteIP $s.RemoteEndpoint
            if ("$($s.SessionId) $ip $($s.TlsInfo.Protocol) $($s.TlsInfo.CertSubject)".ToLower() -notlike "*$fl*") { continue }
            $tlsLabel = if ($s.TlsInfo.Used) { "TLS: $($s.TlsInfo.Protocol)  $($s.TlsInfo.Cipher)/$($s.TlsInfo.CipherBits)-bit" } else { 'No TLS' }
            $sn = [System.Windows.Forms.TreeNode]::new("$($s.SessionId)  [$ip]  [$tlsLabel]")
            $sn.Tag = @{ Type='Session'; Session=$s }
            $sn.ForeColor = if ($s.TlsInfo.Used) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DimGray }
            [void]$tlsTree.Nodes.Add($sn)
        }
        $tlsTree.EndUpdate()
    }.GetNewClosure())

    $tlsTree.Add_AfterSelect({
        $node = $Global:UI_tlsTree.SelectedNode
        if ($null -eq $node -or $null -eq $node.Tag) { return }
        if ($null -ne $Global:UI_tabs -and $Global:UI_tabs.SelectedTab -ne $Global:UI_tabProto) { $Global:UI_tabs.SelectedTab = $Global:UI_tabProto }
        $tag = $node.Tag
        switch ($tag.Type) {
            'Session' { PopulateGridSession $Global:UI_grid $tag.Session; if ($null -ne $Global:UI_detailRtb) { $Global:UI_detailRtb.Clear() } }
            'Mail'    { PopulateGridMail    $Global:UI_grid $tag.Session $tag.Mail; if ($null -ne $Global:UI_detailRtb) { $Global:UI_detailRtb.Clear() } }
        }
    })

    # Grid row -> detail
    $Global:grid.Add_SelectionChanged({
        if ($Global:grid.SelectedRows.Count -eq 0) { return }
        $row = $Global:grid.SelectedRows[0]
        if ($Global:grid.Columns.Contains('Event') -and $Global:grid.Columns.Contains('Data')) {
            $entry = @{
                SequenceNumber = $row.Cells['Seq#'].Value
                DateTime       = $row.Cells['Time'].Value
                Event          = $row.Cells['Event'].Value
                Data           = $row.Cells['Data'].Value
                Context        = if ($Global:grid.Columns.Contains('Context')) { $row.Cells['Context'].Value } else { '' }
            }
            Update-DetailPanel $detailRtb $entry
        }
    }.GetNewClosure())

    # Tab switch -> lazy-load stats / errors
    $tabs.Add_SelectedIndexChanged({
        if ($Global:Sessions.Count -eq 0) { return }
        if ($tabs.SelectedTab -eq $tabStats -and $statsPanel.Controls.Count -eq 0) {
            $stats = Get-Statistics $Global:Sessions
            Build-StatsTab $statsPanel $stats
        }
        if ($tabs.SelectedTab -eq $tabErrors -and $errorsPanel.Controls.Count -eq 0) {
            Build-ErrorsTab $errorsPanel $Global:Sessions
        }
    }.GetNewClosure())

    # ---- Open Files ----
    $doOpen = {
        $ofd = [System.Windows.Forms.OpenFileDialog]::new()
        $ofd.Title = 'Select Exchange SMTP Log File(s)'
        $ofd.Filter = 'Log Files (*.log;*.csv)|*.log;*.csv|All Files (*.*)|*.*'
        $ofd.Multiselect = $true
        if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $files = $ofd.FileNames
        if ($files.Count -eq 0) { return }

        # Validate each selected file before touching any UI state
        $validFiles  = [System.Collections.Generic.List[string]]::new()
        $badReasons  = [System.Collections.Generic.List[string]]::new()
        $warnMsgs    = [System.Collections.Generic.List[string]]::new()
        foreach ($f in $files) {
            $result = Test-LogFile $f
            if (-not $result.OK) {
                $badReasons.Add($result.Reason)
                WriteAppLog 'WARN' "Skipping file: $($result.Reason)"
            } else {
                $validFiles.Add($f)
                if ($result.Warning -ne '') {
                    $warnMsgs.Add($result.Warning)
                    WriteAppLog 'WARN' $result.Warning
                }
            }
        }

        if ($badReasons.Count -gt 0) {
            $msg = "The following file(s) cannot be loaded and will be skipped:`n`n" + ($badReasons -join "`n")
            if ($validFiles.Count -gt 0) { $msg += "`n`nContinue with the remaining $($validFiles.Count) valid file(s)?" }
            $buttons = if ($validFiles.Count -gt 0) { [System.Windows.Forms.MessageBoxButtons]::YesNo } else { [System.Windows.Forms.MessageBoxButtons]::OK }
            $answer  = [System.Windows.Forms.MessageBox]::Show($msg, 'File Validation', $buttons, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($validFiles.Count -eq 0 -or $answer -eq [System.Windows.Forms.DialogResult]::No) { return }
        }

        if ($warnMsgs.Count -gt 0) {
            $msg = "Note:`n`n" + ($warnMsgs -join "`n") + "`n`nThe file(s) will still be loaded."
            [void][System.Windows.Forms.MessageBox]::Show($msg, 'File Warning', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }

        $files = $validFiles.ToArray()

        $Global:Sessions = [System.Collections.Hashtable]::Synchronized(@{})
        $grid.Rows.Clear(); $grid.Columns.Clear(); $detailRtb.Clear()
        $statsPanel.Controls.Clear(); $errorsPanel.Controls.Clear()
        $searchGrid.Rows.Clear(); $searchGrid.Columns.Clear()
        $Global:treeView.Nodes.Clear(); $ehloTree.Nodes.Clear(); $tlsTree.Nodes.Clear()
        $tbExport.Enabled = $false; $miExport.Enabled = $false
        $tbProg.Style   = [System.Windows.Forms.ProgressBarStyle]::Blocks
        $tbProg.Minimum = 0; $tbProg.Maximum = 100; $tbProg.Value = 0
        $tbProg.Width = 200
        $tbLabel.Text   = 'Preparing...'
        $sbMsg.Text     = "Preparing parse plan for $($files.Count) file(s)..."
        $sbFiles.Text   = "Files: $($files.Count)"
        $sbTime.Text    = 'Remaining: calculating'
        $form.Cursor    = [System.Windows.Forms.Cursors]::WaitCursor
        $miOpen.Enabled = $false; $tbOpen.Enabled = $false
        [System.Windows.Forms.Application]::DoEvents()

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $parseState = New-ParseOperationState $files
        $tbLabel.Text = 'Counting lines...'
        $sbMsg.Text   = "Preparing $($parseState.FileTotal) file(s)..."
        [System.Windows.Forms.Application]::DoEvents()

        # Store all controls tick handler needs in globals - nested closures are unreliable in PS 5.1
        $Global:UI_tbProg     = $tbProg
        $Global:UI_tbLabel    = $tbLabel
        $Global:UI_sbMsg      = $sbMsg
        $Global:UI_sbFiles    = $sbFiles
        $Global:UI_sbTime     = $sbTime
        $Global:UI_sbSessions = $sbSessions
        $Global:UI_form       = $form
        $Global:UI_treeView   = $Global:treeView
        $Global:UI_ehloTree   = $ehloTree
        $Global:UI_tlsTree    = $tlsTree
        $Global:UI_grid       = $Global:grid
        $Global:UI_detailRtb  = $detailRtb
        $Global:UI_tabs       = $tabs
        $Global:UI_tabProto   = $tabProto
        $Global:UI_miOpen     = $miOpen
        $Global:UI_tbOpen     = $tbOpen
        $Global:UI_tbExport   = $tbExport
        $Global:UI_miExport   = $miExport
        $Global:UI_files      = $files
        $Global:UI_sw         = $sw

        $parseTimer = [System.Windows.Forms.Timer]::new()
        $parseTimer.Interval = 50
        $Global:ActiveParseState = $parseState
        $Global:ActiveParseTimer = $parseTimer

        $parseTimer.Add_Tick({
            try {
                Invoke-ParseBatch -State $Global:ActiveParseState -BatchSize 2000

                $pct = if ($Global:ActiveParseState.TotalLines -gt 0) {
                    [int](($Global:ActiveParseState.ProcessedLines * 100) / $Global:ActiveParseState.TotalLines)
                } else { 0 }
                $Global:UI_tbProg.Value = [Math]::Min(100, [Math]::Max(0, $pct))
                $fi = [Math]::Min($Global:ActiveParseState.FileIndex + 1, $Global:ActiveParseState.FileTotal)
                $Global:UI_tbLabel.Text = "$($Global:ActiveParseState.Phase) - file $fi/$($Global:ActiveParseState.FileTotal): $($Global:ActiveParseState.CurrentFileName)  ${pct}%"
                $Global:UI_sbMsg.Text   = "File $fi/$($Global:ActiveParseState.FileTotal) | Line $($Global:ActiveParseState.CurrentFileLines)/$($Global:ActiveParseState.CurrentFileTotal) | Total: $($Global:ActiveParseState.ProcessedLines)/$($Global:ActiveParseState.TotalLines)"
                $Global:UI_sbFiles.Text = "Files: $fi / $($Global:ActiveParseState.FileTotal)"
                $Global:UI_sbTime.Text  = "Remaining: $($Global:ActiveParseState.RemainingLines) lines"

                if ($Global:ActiveParseState.IsComplete) {
                    $Global:ActiveParseTimer.Stop()
                    $Global:ActiveParseTimer.Dispose()
                    $Global:UI_sw.Stop()
                    $Global:UI_form.Cursor      = [System.Windows.Forms.Cursors]::Default
                    $Global:UI_tbProg.Width     = 0
                    $Global:UI_miOpen.Enabled   = $true; $Global:UI_tbOpen.Enabled   = $true
                    $Global:Sessions            = $Global:ActiveParseState.Sessions
                    $cnt = $Global:Sessions.Count
                    PopulateTreeView  $Global:UI_treeView $Global:Sessions
                    PopulateEhloTree  $Global:UI_ehloTree $Global:Sessions
                    PopulateTlsTree   $Global:UI_tlsTree  $Global:Sessions
                    PopulateGridConnector $Global:UI_grid $Global:Sessions $null
                    $Global:UI_tbExport.Enabled = $true; $Global:UI_miExport.Enabled = $true
                    $elapsed = $Global:UI_sw.Elapsed.ToString('mm\:ss\.fff')
                    $Global:UI_tbLabel.Text    = "Done. $cnt sessions."
                    $Global:UI_sbMsg.Text      = "Parsed $($Global:UI_files.Count) file(s) - $cnt sessions loaded"
                    $Global:UI_sbSessions.Text = "Sessions: $cnt"
                    $Global:UI_sbTime.Text     = "Parse time: $elapsed"
                    WriteAppLog 'INFO' "Parse complete. Sessions: $cnt  Time: $elapsed"
                    $Global:ActiveParseTimer = $null
                    $Global:ActiveParseState = $null
                }
            } catch {
                $Global:ActiveParseTimer.Stop()
                $Global:ActiveParseTimer.Dispose()
                $Global:UI_sw.Stop()
                $Global:UI_form.Cursor    = [System.Windows.Forms.Cursors]::Default
                $Global:UI_tbProg.Width   = 0
                $Global:UI_miOpen.Enabled = $true; $Global:UI_tbOpen.Enabled = $true
                WriteAppLog 'ERROR' "Parse error: $_"
                [void][System.Windows.Forms.MessageBox]::Show("Parse error:`n$($_.Exception.Message)", 'Error',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
                $Global:UI_sbMsg.Text    = 'Parse failed.'
                $Global:ActiveParseTimer = $null
                $Global:ActiveParseState = $null
            }
        })

        $parseTimer.Start()
    }.GetNewClosure()
    $miOpen.Add_Click($doOpen); $tbOpen.Add_Click($doOpen)

    # ---- Export HTML ----
    $doExport = {
        if ($Global:Sessions.Count -eq 0) { return }
        $sfd = [System.Windows.Forms.SaveFileDialog]::new()
        $sfd.Title = 'Export HTML Report'; $sfd.Filter = 'HTML Files (*.html)|*.html'
        $sfd.FileName = "SMTPLogReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $form.Cursor  = [System.Windows.Forms.Cursors]::WaitCursor
        $tbLabel.Text = 'Exporting...'
        try {
            $stats = Get-Statistics $Global:Sessions
            Export-HtmlReport -OutputPath $sfd.FileName -Sessions $Global:Sessions -Stats $stats
            $tbLabel.Text = 'Export complete.'
            WriteAppLog 'INFO' "HTML exported: $($sfd.FileName)"
            [void][System.Windows.Forms.MessageBox]::Show("Report saved:`n$($sfd.FileName)", 'Export Complete',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            WriteAppLog 'ERROR' "Export failed: $_"
            [void][System.Windows.Forms.MessageBox]::Show("Export failed:`n$_", 'Error',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }.GetNewClosure()
    $miExport.Add_Click($doExport); $tbExport.Add_Click($doExport)

    # ---- View Log ----
    $doLog = {
        if (Test-Path $Global:LogPath) { Start-Process 'notepad.exe' $Global:LogPath }
        else {
            [void][System.Windows.Forms.MessageBox]::Show('No log file exists yet.', 'Log',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
    $miViewLog.Add_Click($doLog); $tbLog.Add_Click($doLog)

    # ---- Exit / About ----
    $miExit.Add_Click({ $form.Close() }.GetNewClosure())
    $miAbout.Add_Click({
        $dlg = [System.Windows.Forms.Form]::new()
        $dlg.Text = 'About SMTP Protocol Log Parser'
        $dlg.Size = [System.Drawing.Size]::new(420, 290)
        $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
        $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false

        $ab_bmp = Get-LogoBitmap
        $ab_h   = 56; $ab_w = [int]($ab_bmp.Width * $ab_h / $ab_bmp.Height)
        $ab_scaled = [System.Drawing.Bitmap]::new($ab_bmp, [System.Drawing.Size]::new($ab_w, $ab_h))
        $ab_bmp.Dispose()
        $ab_pb  = [System.Windows.Forms.PictureBox]::new()
        $ab_pb.Image    = $ab_scaled
        $ab_pb.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::AutoSize
        $ab_pb.Location = [System.Drawing.Point]::new(20, 18)

        $ab_l1  = [System.Windows.Forms.Label]::new()
        $ab_l1.Text     = 'SMTP Protocol Log Parser v1.0'
        $ab_l1.Font     = [System.Drawing.Font]::new('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
        $ab_l1.Location = [System.Drawing.Point]::new(20, 86)
        $ab_l1.AutoSize = $true

        $ab_l2  = [System.Windows.Forms.Label]::new()
        $ab_l2.Text     = "Exchange Server SMTP Receive Log Analysis Tool`nPowerShell 5.1 Compatible  -  No external dependencies"
        $ab_l2.Location = [System.Drawing.Point]::new(20, 116)
        $ab_l2.AutoSize = $true

        $ab_link = [System.Windows.Forms.LinkLabel]::new()
        $ab_link.Text     = 'www.cloudvision.com.tr'
        $ab_link.Location = [System.Drawing.Point]::new(20, 165)
        $ab_link.AutoSize = $true
        $ab_link.Add_LinkClicked({ Start-Process 'https://www.cloudvision.com.tr' })

        $ab_ok  = [System.Windows.Forms.Button]::new()
        $ab_ok.Text          = 'OK'
        $ab_ok.DialogResult  = [System.Windows.Forms.DialogResult]::OK
        $ab_ok.Location      = [System.Drawing.Point]::new(160, 205)
        $ab_ok.Width         = 90
        $dlg.AcceptButton    = $ab_ok

        [void]$dlg.Controls.AddRange(@($ab_pb, $ab_l1, $ab_l2, $ab_link, $ab_ok))
        [void]$dlg.ShowDialog($form)
        $ab_scaled.Dispose(); $dlg.Dispose()
    }.GetNewClosure())

    # ---- Search ----
    $doSearch = {
        if ($Global:Sessions.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('No data loaded. Open a log file first.', 'Search',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $fIP   = $txtSenderIP.Text.Trim().ToLower()
        $fAddr = $txtSenderAddr.Text.Trim().ToLower()
        $fRcpt = $txtRecipient.Text.Trim().ToLower()
        $fSID  = $txtSessionID.Text.Trim().ToLower()

        $results = $Global:Sessions.Values | Where-Object {
            ($fIP   -eq '' -or (Get-RemoteIP $_.RemoteEndpoint).ToLower() -like "*$fIP*") -and
            ($fAddr -eq '' -or $_.SenderAddress.ToLower() -like "*$fAddr*") -and
            ($fRcpt -eq '' -or ($_.Recipients -join ' ').ToLower() -like "*$fRcpt*") -and
            ($fSID  -eq '' -or $_.SessionId.ToLower() -like "*$fSID*")
        }

        $searchGrid.Rows.Clear()
        Set-GridCols $searchGrid `
            @('Session-ID','Remote IP','Start Time','End Time','Sender','Recipients','Status','Error','Error Message') `
            @(140,110,145,145,160,200,75,65,200)
        foreach ($s in ($results | Sort-Object { $_.StartTime })) {
            $ri = $searchGrid.Rows.Add($s.SessionId, (Get-RemoteIP $s.RemoteEndpoint), $s.StartTime,
                  $s.EndTime, $s.SenderAddress, ($s.Recipients -join '; '),
                  $s.Status, $s.ErrorCode, $s.ErrorMessage)
            $searchGrid.Rows[$ri].Tag = $s
            $searchGrid.Rows[$ri].DefaultCellStyle.BackColor = switch ($s.Status) {
                'Error'      { [System.Drawing.Color]::FromArgb(255,235,235) }
                'Incomplete' { [System.Drawing.Color]::FromArgb(255,250,225) }
                default      { [System.Drawing.Color]::White }
            }
        }
        $sbMsg.Text = "Search: $($searchGrid.Rows.Count) result(s)"
    }.GetNewClosure()
    $btnSearch.Add_Click($doSearch)

    # Enter key in search fields triggers search
    foreach ($tb in @($txtSenderIP,$txtSenderAddr,$txtRecipient,$txtSessionID)) {
        $tb.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Return) { & $doSearch }
        }.GetNewClosure())
    }

    $btnClear.Add_Click({
        $txtSenderIP.Clear(); $txtSenderAddr.Clear(); $txtRecipient.Clear(); $txtSessionID.Clear()
        $searchGrid.Rows.Clear(); $searchGrid.Columns.Clear()
        $sbMsg.Text = 'Search cleared'
    }.GetNewClosure())

    # Double-click search result -> navigate tree
    $searchGrid.Add_CellDoubleClick({
        if ($searchGrid.SelectedRows.Count -eq 0) { return }
        $target = $searchGrid.SelectedRows[0].Tag
        if ($null -eq $target) { return }
        foreach ($cn in $Global:treeView.Nodes) {
            foreach ($sn in $cn.Nodes) {
                $t = $sn.Tag
                if ($null -ne $t -and $t.Type -eq 'Session' -and $t.Session.SessionId -eq $target.SessionId) {
                    $leftTabs.SelectedTab = $tabSessions.Page
                    $Global:treeView.SelectedNode = $sn; $sn.EnsureVisible()
                    $tabs.SelectedTab = $tabProto
                    return
                }
            }
        }
    }.GetNewClosure())

    return $form
}

# ================================================================
#  ENTRY POINT
# ================================================================
WriteAppLog 'INFO' 'Application started'
$form = Build-MainForm
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
WriteAppLog 'INFO' 'Application closed'
