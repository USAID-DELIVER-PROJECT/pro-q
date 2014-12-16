<?xml version="1.0"?>
<?xml-stylesheet type="text/xsl" href="Answers.xsl"?>

<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">

<xsl:template match="/">
  <HTML>	
    <BODY>

	<table border="0" width="100%" cellspacing="0">
	  <tr>
    	    <td width="70%"></td>
	    <xsl:apply-templates select="Report/Header/Aggregation" />		    
	  </tr>
	  <tr>
	    <td width="70%"></td>
	    <xsl:apply-templates select="Report/Header/Program" />		    
	  </tr>
	  <tr>
	    <td width="70%"></td>
	    <xsl:apply-templates select="Report/Header/Use" />		    
	  </tr>
	</table>

        <xsl:apply-templates select="Report/Header/Report_Title" />		    
	

	<TABLE border="1" cellspacing="0" width="100%">

	<TR>
	<TD bgcolor="#C0C0C0"> <b><u>Question</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Answer</u></b> </TD>
	</TR>

	<xsl:for-each select="Report/Responses/Response">
	  <TR>
          <xsl:apply-templates select="Question" />
          <xsl:apply-templates select="Answer" />
	  </TR>
	</xsl:for-each> 
	</TABLE>

	<p></p>
        <xsl:apply-templates select="Report/Other/LabelA" />
	<p></p>
        <xsl:apply-templates select="Report/Other/LabelB" />	

    </BODY>
  </HTML>
</xsl:template>

<xsl:template match="Question">
  <TD width="60%" >
	<b> <xsl:value-of /> </b>
  </TD>
</xsl:template>

<xsl:template match="Answer">
	<TD width="40%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>


<xsl:template match="Report/Header/Aggregation">
	<td width="30%"><b><xsl:value-of /></b></td>	
</xsl:template>

<xsl:template match="Report/Header/Program">
	<td width="30%"><i><xsl:value-of /></i></td>	
</xsl:template>

<xsl:template match="Report/Header/Use">
	<td width="30%"><i><xsl:value-of /></i></td>	
</xsl:template>

<xsl:template match="Report/Header/Report_Title">
	<H2><xsl:value-of /></H2>	
</xsl:template>

<xsl:template match="Report/Other/LabelA">
	<b><xsl:value-of /></b>
</xsl:template>

<xsl:template match="Report/Other/LabelB">
	<b><xsl:value-of /></b>	
</xsl:template>



</xsl:stylesheet>