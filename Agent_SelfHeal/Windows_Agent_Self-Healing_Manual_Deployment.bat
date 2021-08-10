@ECHO OFF
REM ============================================
REM WINDOWS AGENT SELF-HEALING MANUAL DEPLOYMENT
REM ============================================
REM.
IF NOT EXIST C:\IT MD C:\IT
IF NOT EXIST C:\IT\Scripts MD C:\IT\Scripts
IF NOT EXIST C:\IT\Scripts\Tasks MD C:\IT\Scripts\Tasks
CURL https://raw.githubusercontent.com/CW-Khristos/scripts/master/Agent_SelfHeal/AgentSelfHeal.ps1 -o C:\IT\Scripts\AgentSelfHeal.ps1
CURL https://raw.githubusercontent.com/CW-Khristos/scripts/master/Agent_SelfHeal/N-able%20Windows%20Agent%20Self-Healing.xml -o "C:\IT\Scripts\Tasks\N-able Windows Agent Self-Healing.xml"
SchTasks /CREATE /TN "IPM Computers\N-able Windows Agent Self-Healing" /XML "C:\IT\Scripts\Tasks\N-able Windows Agent Self-Healing.xml" /F
PowerShell -ExecutionPolicy Bypass -File C:\IT\Scripts\AgentSelfHeal.ps1
