#pragma once

namespace FindFrameAlgorithms
{
	class FrameVisitor
	{
	public:
		virtual ~FrameVisitor() = 0
		{
		}
		virtual BOOL VisitFrame(CComQIPtr<IHTMLWindow2> spFrame) = 0;
	};
}
