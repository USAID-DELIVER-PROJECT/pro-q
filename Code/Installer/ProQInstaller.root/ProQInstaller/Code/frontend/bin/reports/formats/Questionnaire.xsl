<?xml version="1.0"?>
<?xml-stylesheet type="text/xsl" href="Questionnaire.xsl"?>

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
	

	<xsl:for-each select="Report/Use">	
          <xsl:apply-templates select="UseName" />
        
          <TABLE border="1" cellspacing="0" width="100%">

  	    <TR>
	      <TD bgcolor="#C0C0C0"> <b><u>Question</u></b> </TD>
	      <TD align = "center" bgcolor="#C0C0C0"> <b><u>Answer</u></b> </TD>
	    </TR>

 	    <xsl:for-each select="Category">
	      <TR>
                <xsl:apply-templates select="CategoryName" />
	      </TR>
          
    	      <xsl:for-each select="Topic">
	        <TR>
                  <xsl:apply-templates select="TopicName" />   	  
                </TR>

     	        <xsl:for-each select="SubTopic">
  	          <TR>
                    <xsl:apply-templates select="SubTopicName" />   	  
                  </TR>

     	          <xsl:for-each select="Response">            
                    <TR>    
		      <td width="80%">	
                      <xsl:apply-templates select="Question" />   	  
		      </td>	
                      <xsl:apply-templates select="Answer" />
                    </TR>   	
     	          </xsl:for-each> 
		</xsl:for-each>		

     	        <xsl:for-each select="Response">            
                  <TR>    
                    <td width="80%">
                    <xsl:apply-templates select="BrandName" />
                    <xsl:apply-templates select="Question" />   	  
                    </td>
                    <xsl:apply-templates select="Answer" />
                  </TR>   	
     	        </xsl:for-each> 

   	      </xsl:for-each> 
	    </xsl:for-each> 
 	  </TABLE>	
	<p></p>
        </xsl:for-each>        

    </BODY>
  </HTML>
</xsl:template>

<xsl:template match="UseName">
	<H3> <xsl:value-of /> </H3>  
</xsl:template>

<xsl:template match="CategoryName">
  <TD width="100%" colspan="2" bgcolor="#FFFF00">
	<H3> <xsl:value-of /> </H3>
  </TD>
</xsl:template>

<xsl:template match="TopicName">
  <TD width="100%" colspan="2">
	<b> <xsl:value-of /> </b>
  </TD>
</xsl:template>

<xsl:template match="SubTopicName">
  <TD width="100%" colspan="2">
	<u><xsl:value-of /> </u>
  </TD>
</xsl:template>


<xsl:template match="BrandName">
	<u><xsl:value-of />&#160;</u>	
</xsl:template>

<xsl:template match="Question">
	<xsl:value-of />&#160;	
</xsl:template>

<xsl:template match="Answer">
	<td width="20%"><xsl:value-of /> &#160; </td>	
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


</xsl:stylesheet>