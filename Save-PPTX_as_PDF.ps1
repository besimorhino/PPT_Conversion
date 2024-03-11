# By MBD 20240302

# based on:
# https://stackoverflow.com/questions/73572064/converting-pptx-file-to-pdf-using-powershell
# https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.visible


# get the dir to convert
param ($Path) 

Get-ChildItem -Path $path -Include "*.ppt", "*.pptx" -Recurse | 
ForEach-Object {

    # get file name without the extention, and add on ".pdf"
    $pdf = ($_.Name.Split(".")[0]) + ".pdf"

    # get the full path of current PPT or PPTX and add a PDFs subdirectory
    $Target_Dir = $_.DirectoryName + "\PDFs"

    # test to see if the target_dir exists. if not, create one.  
    if (!(Test-Path $Target_Dir -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $Target_Dir
    }

    # create the full output file
    $OutPut_File = $Target_Dir + "\" + $pdf

    # launch PowerPoint
    $ppt = New-Object -ComObject PowerPoint.Application

    # This was in the stack overflow article... but this causes an error on my box.
    # $ppt.Visible = True
    # using msoTrue gets rid of the error
    $ppt.Visible = "msoTrue"

    # open the PPT or PPTX
    $presentation = $ppt.Presentations.Open($_)
    "Converting $_"
  
    # yay! the reason d'etre for this script. Save as PDF, please and thank you.
    $presentation.SaveAs($OutPut_File, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)

    # sometimes the close method gives a warning. This one second sleep removes the errors on my system.
    sleep(1)
    $presentation.Close()
    $ppt.Quit()
}

