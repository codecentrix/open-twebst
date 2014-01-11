<xsl:stylesheet version = '1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'>
	<xsl:template match="topic">
		<html>
			<head>
				<title>
					<xsl:apply-templates select="title"/>
				</title>

				<link href="../CSS/help.css" rel="stylesheet" type="text/css" />
			</head>

			<body>
				<div class="topicHighlight">
					<h2>
						<xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text>
						<xsl:value-of select="@name"/>
						<xsl:text> - web automation </xsl:text>
						<xsl:value-of select="@type"/>
					</h2>
				</div>

				<h1><xsl:apply-templates select="description"/></h1>

				<div class="section">Applies to:</div>
				<dl>
					<dd>
						<xsl:apply-templates select="apply"/>
					</dd>

					<dd>
						<code><xsl:text disable-output-escaping="yes">&amp;nbsp;</xsl:text><xsl:value-of select="call"/></code>
					</dd>
				</dl>

				<div class="section">Arguments:</div>
				<dl>
					<xsl:apply-templates select="arguments"/>
				</dl>

				<div class="section">Remarks:</div>
				<dl>
					<dd>
						<xsl:apply-templates select="remarks"/>
					</dd>
				</dl>

				<div class="section">Example:</div>
				<dl><dd><xsl:apply-templates select="example"/></dd></dl>

				<div class="section">See also:</div>
				<dl>
					<dd><xsl:apply-templates select="seealso"/></dd>
				</dl>

				<hr SIZE="1"></hr>
				<small><i>
					<a target="_blank" href="http://www.codecentrix.com">
						<xsl:text disable-output-escaping="yes">&amp;copy;</xsl:text>
						2014 CodeCentrix Software. All rights reserved.
					</a>
				</i></small>
			</body>
		</html>
	</xsl:template>

	<xsl:template match="apply_item">
		<xsl:apply-templates />
		<xsl:if test="position()!=last()">, </xsl:if>
		<!--xsl:if test="position()=last()">.</xsl:if-->
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

	<xsl:template match="title|description|arg_summary|apply|remarks|example|seealso">
		<xsl:apply-templates />
	</xsl:template>
	
	<xsl:template match="arguments">
		<dd>
			<xsl:apply-templates select="arg_summary"/>
			<ul>
				<xsl:apply-templates select="arg"/>
			</ul>
		</dd>
	</xsl:template>

	<xsl:template match="arg">
		<li>
			<em><xsl:value-of select="@name"/><![CDATA[  ]]></em>
			<xsl:apply-templates />
		</li>
	</xsl:template>

	<xsl:template match="jscode">
			<pre><code><xsl:apply-templates /></code></pre></xsl:template>

	<xsl:template match="jscomment">
			<span STYLE="color: green"><xsl:value-of select="."/></span>
	</xsl:template>

	<xsl:template match="seealso_item">
		<xsl:apply-templates />
		<xsl:if test="position()!=last()"> | </xsl:if>
	</xsl:template>

	<xsl:template match="keyword">
		<span STYLE="color: blue"><xsl:value-of select="."/></span>
	</xsl:template>
	
	<xsl:template match="string">
		<span STYLE="color: red"><xsl:value-of select="."/></span>
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
	
	<xsl:template match="*">
		<xsl:element name="{name()}">
			<xsl:apply-templates />
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
