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

        <xsl:apply-templates select="Report/Header/Report_Title" />		           <xsl:apply-templates select="Report/Header/View_By" />		    


	<TABLE border="1" cellspacing="0" >
  	  <TR>
	    <xsl:for-each select="Report/Columns/Column">
	      <TD width="150" align = "center" bgcolor="#C0C0C0">
                <xsl:apply-templates select="ColHeading" />
	      </TD>
	    </xsl:for-each> 
	  </TR>

          <xsl:for-each select="Report/Records/Record">
   	    <TR>
 	      <xsl:for-each select="Field">
                  <xsl:apply-templates select="Value" />
	      </xsl:for-each> 	
 	    </TR>
	  </xsl:for-each> 
	</TABLE>
    </BODY>
  </HTML>
</xsl:template>


<xsl:template match="Value">
	<TD align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="ColHeading">
	<b>&#160; <xsl:value-of /> </b>
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

<xsl:template match="Report/Header/View_By">
	<H3><xsl:value-of />&#160; </H3>	
</xsl:template>


</xsl:stylesheet>