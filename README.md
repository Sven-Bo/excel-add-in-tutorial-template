# How to Create a Custom Excel Add-in (Step-by-Step Guide)
Are you tired of performing repetitive tasks in Microsoft Excel? Do you want to add new functionality to Excel and make it even more powerful? Then this tutorial is for you! In this step-by-step guide, you'll learn how to create a custom add-in for Excel using built-in tools. Whether you're a beginner or an experienced Excel user, this tutorial is easy to follow and packed with helpful tips and tricks.

## What You'll Learn

- How to create a new add-in in Excel
- How to add custom macros to your add-in
- How to customize the ribbon and toolbar for easy access to your add-in
- How to package and distribute your add-in to others

By the end of this tutorial, you'll have a custom Excel add-in that can help boost your productivity and make your work in Excel more efficient.


## Video Tutorial
[![YouTube Video](https://img.youtube.com/vi/avdVI14AxzM/0.jpg)](https://youtu.be/avdVI14AxzM)

## Reference
The following website will be used as a reference in this tutorial:<br/>
https://bettersolutions.com/vba/ribbon/custom-ui-editor.htm

## List of imageMSO values and associated pictures
https://bert-toolkit.com/imagemso-list.html

## XML Code for the UI
```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"> 
  <ribbon startFromScratch="false"> 
    <tabs> 
      <tab id="CustomTab" label="My Add-in"> 
        <group id="SimpleControls" label="ChatGPT"> 
          <button id="Button1" image="ai" size="large" 
                  label="Ask AI" 
                  screentip="Ask AI" 
                  onAction="OpenAI_Completion"/> 
        </group>
        <group id="Utils" label="Utils"> 
          <button id="Button2" imageMso="EqualSign" size="normal" 
                  label="IfErrorBlank" 
                  onAction="IfErrorBlank"/> 
          <button id="Button3" imageMso="EqualSign" size="normal" 
                  label="IfErrorZero" 
                  onAction="IfErrorZero"/> 
        </group>  
      </tab> 
    </tabs> 
  </ribbon> 
</customUI> 
```


## About Me
- Sven from Coding Is Fun
- YouTube: https://youtube.com/c/CodingIsFun
- Website: https://pythonandvba.com

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/X7X47Q0EG)

## Feedback
If you have any feedback, please reach out to me at contact@pythonandvba.com
![Logo](https://www.pythonandvba.com/banner-img)
