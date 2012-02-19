#pragma once
#include "Common.h"
#include "FrameVisitor.h"
#include "FindFrame.h"

namespace FindFrameAlgorithms
{
	class SearchFrameVisitor : public FindFrameAlgorithms::FrameVisitor
	{
	public:
		SearchFrameVisitor(SAFEARRAY* psa);
		~SearchFrameVisitor(void);
		virtual BOOL VisitFrame(CComQIPtr<IHTMLWindow2> spElement);
		CComQIPtr<IHTMLWindow2> GetFoundHtmlFrame();

	private:
		BOOL GetFrameAttributeValue(CComQIPtr<IHTMLWindow2> spFrame, const String& sAttributeName, String* pAttributeValue);

	private:
		CComQIPtr<IHTMLWindow2>    m_spFoundFrame;
		std::list<DescriptorToken> m_tokens;
		int                        m_nSearchIndex;
	};


	class SearchFrameCollectionVisitor : public FindFrameAlgorithms::FrameVisitor
	{
	public:
		SearchFrameCollectionVisitor(SAFEARRAY* psa, LONG nSearchFlags);
		virtual BOOL VisitFrame(CComQIPtr<IHTMLWindow2> spFrame);
		const list<CAdapt<CComQIPtr<IHTMLWindow2> > >& GetFoundHtmlFrameCollection();

	private:
		SearchFrameVisitor                      m_frameVisitor;
		list<CAdapt<CComQIPtr<IHTMLWindow2> > > m_foundFrameCollection;
	};
}
