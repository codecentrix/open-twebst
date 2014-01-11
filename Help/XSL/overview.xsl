<xsl:stylesheet version = '1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'>
	<xsl:template match="overview">
		<html>
			<head>
				<xsl:apply-templates select="metadesc"/>
				<xsl:apply-templates select="metakeywords"/>

				<title>
					<xsl:apply-templates select="title"/>
				</title>

				<link href="./CSS/help.css" rel="stylesheet" type="text/css"/>
			</head>

			<body>
				<div class="topicHighlight">
					<h2>
						<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
						<xsl:value-of select="@name"/>
						<xsl:text> - IE macro </xsl:text>
						<xsl:value-of select="@type"/>
					</h2>
				</div>

				<dl>
					<dd>
						<p>
							<xsl:apply-templates select="description"/>
						</p>
					</dd>
					<xsl:apply-templates select="creation"/>
				</dl>

				<xsl:apply-templates select="methods"/>
				<xsl:apply-templates select="properties"/>
				<xsl:apply-templates select="constants"/>

				<div class="section">See also:</div>
				<dl>
					<dd>
						<xsl:apply-templates select="seealso/seealso_item"/>
					</dd>
				</dl>

				<hr SIZE="1"></hr>
				<small>
					<i>
						<a target="_blank" href="http://www.codecentrix.com">
							<xsl:text disable-output-escaping="yes">&amp;copy;</xsl:text>
							2014 CodeCentrix Software. All rights reserved.
						</a>
					</i>
				</small>

				<!-- Start of StatCounter Code -->
				<script type="text/javascript">
				sc_project=1010390; 
				sc_invisible=1; 
				sc_partition=9; 
				sc_security="a4cc6c6f"; 
				</script>

				<script type="text/javascript" src="http://www.statcounter.com/counter/counter.js"></script>
				<noscript><img src="http://c10.statcounter.com/1010390/0/a4cc6c6f/1/" alt="statcounter" /></noscript>
				<!-- End of StatCounter Code -->
			</body>
		</html>
	</xsl:template>


	<xsl:template match="description">
		<xsl:apply-templates />
	</xsl:template>

	<xsl:template match="creation">
		<xsl:apply-templates select="creat_method" />
	</xsl:template>

	<xsl:template match="creat_method">
		<dd>
			<code>
				<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
				<xsl:apply-templates />
			</code>
		</dd>
	</xsl:template>

	<xsl:template match="methods">
		<div class="section">Methods</div>
		<dl>
			<dd>
				<table cellspacing="0" border="1" width="90%">
					<xsl:apply-templates select="method"/>
				</table>
			</dd>
		</dl>
	</xsl:template>

	<xsl:template match="method">
		<tr valign="top">
			<xsl:apply-templates select="method_name"/>
			<xsl:apply-templates select="method_description"/>
		</tr>
	</xsl:template>

	<xsl:template match="method_name">
		<td width="20%"><xsl:apply-templates /></td>
	</xsl:template>

	<xsl:template match="method_description">
		<td width="80%"><xsl:apply-templates /></td>
	</xsl:template>

	<xsl:template match="properties">
		<div class="section">Properties</div>
		<dl>
			<dd>
				<table cellspacing="0" border="1" width="90%">
					<xsl:apply-templates select="property"/>
				</table>
			</dd>
		</dl>
	</xsl:template>

	<xsl:template match="property">
		<tr valign="top">
			<xsl:apply-templates select="prop_name"/>
			<xsl:apply-templates select="prop_desc"/>
		</tr>
	</xsl:template>

	<xsl:template match="prop_name">
		<td width="20%"><xsl:apply-templates /></td>
	</xsl:template>

	<xsl:template match="prop_desc">
		<td width="80%"><xsl:apply-templates /></td>
	</xsl:template>

	<xsl:template match="constants">
		<a name="#constants"/>
		<div class="section">Constants</div>
		<dl>
			<dd>
				<table cellspacing="0" border="1" width="90%">
					<xsl:apply-templates select="constant"/>
				</table>
			</dd>
		</dl>
	</xsl:template>

	<xsl:template match="constant">
		<tr valign="top">
			<xsl:apply-templates select="const_name"/>
			<xsl:apply-templates select="const_desc"/>
		</tr>
	</xsl:template>

	<xsl:template match="const_name">
		<td width="20%"><xsl:apply-templates /></td>
	</xsl:template>

	<xsl:template match="const_desc">
		<td width="80%"><xsl:apply-templates /></td>
	</xsl:template>
	
	<xsl:template match="seealso/seealso_item">
		<xsl:apply-templates />
		<xsl:if test="position()!=last()"> | </xsl:if>
	</xsl:template>

	<xsl:template match="reference">
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="."/>.htm
			</xsl:attribute>

			<xsl:if test="@href!=''">
				<xsl:attribute name="href">
					<xsl:value-of select="@href"/>
				</xsl:attribute>
			</xsl:if>

			<xsl:value-of select="."/>
		</a>
	</xsl:template>

	<xsl:template match="a">
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="@href"/>
			</xsl:attribute>

			<xsl:attribute name="target">
				<xsl:value-of select="@target"/>
			</xsl:attribute>
		
			<xsl:apply-templates />
		</a>
	</xsl:template>
	
	<xsl:template match="metadesc">
		<meta name="Description">
			<xsl:attribute name="content">
				<xsl:apply-templates />
			</xsl:attribute>
		</meta>
	</xsl:template>

	<xsl:template match="metakeywords">
		<meta name="Keywords">
			<xsl:attribute name="content">
				<xsl:apply-templates />
			</xsl:attribute>
		</meta>
	</xsl:template>
	
	<xsl:template match="title">
		<xsl:apply-templates />
	</xsl:template>

	<xsl:template match="*">
		<xsl:element name="{name()}">
			<xsl:apply-templates />
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
