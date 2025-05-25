// xll_template.cpp - Sample xll project.
#include <cmath> // for double tgamma(double)
#include "xll_template.h"

using namespace xll;

AddIn xai_tgamma(
	// Return double, C++ name of function, Excel name.
	Function(XLL_DOUBLE, "xll_tgamma", "TGAMMA")
	// Array of function arguments.
	.Arguments({
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.")
		})
	// Function Wizard help.
	.FunctionHelp("Return the Gamma function value.")
	// Function Wizard category.
	.Category("MATH")
	// URL linked to `Help on this function`.
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
	.Documentation(R"xyzyx(
The <i>Gamma</i> function is \(\Gamma(x) = \int_0^\infty t^{x - 1} e^{-t}\,dt\), \(x \ge 0\).
If \(n\) is a natural number then \(\Gamma(n + 1) = n! = n(n - 1)\cdots 1\).
<p>
Any valid HTML using <a href="https://katex.org/" target="_blank">KaTeX</a> can 
be used for documentation.
)xyzyx")
);
// WINAPI calling convention must be specified
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT // must be specified to export function

	return tgamma(x);
}


// Press Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
// See https://xlladdins.github.io/Excel4Macros/
AddIn xai_macro(
	// C++ function, Excel name of macro
	Macro("xll_macro", "XLL.MACRO")
);
// Macros must have `int WINAPI (*)(void)` signature.
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	// https://xlladdins.github.io/Excel4Macros/reftext.html
	// A1 style instead of default R1C1.
	OPER reftext = Excel(xlfReftext, Excel(xlfActiveCell), OPER(true));
	// UTF-8 strings can be used.
	Excel(xlcAlert, OPER("XLL.MACRO called with : ") & reftext);

	return TRUE;
}

double norm_cdf(double x)
{
	return 0.5 * (1 + std::erf(x / std::sqrt(2.0)));
}
 
double bsm_put(double r, double S, double sigma, double K, double t)
{
	if (sigma <= 0 || t <= 0 || S <= 0 || K <= 0) {
		return std::numeric_limits<double>::quiet_NaN();
	}

	double d1 = (std::log(S / K) + (r + 0.5 * sigma * sigma) * t) / (sigma * std::sqrt(t));
	double d2 = d1 - sigma * std::sqrt(t);

	double put = K * std::exp(-r * t) * norm_cdf(-d2) - S * norm_cdf(-d1);

	return put;
}

AddIn xai_bsm_put(
	// Return double, C++ name of function, Excel name
	Function(XLL_DOUBLE, "xll_bsm_put", "BSM.PUT")
	// Array of function arguments
	.Arguments({
		Arg(XLL_DOUBLE, "r", "is the risk-free interest rate (annualized)."),
		Arg(XLL_DOUBLE, "S", "is the current underlying asset price."),
		Arg(XLL_DOUBLE, "sigma", "is the volatility of returns of the underlying asset (annualized)."),
		Arg(XLL_DOUBLE, "K", "is the strike price."),
		Arg(XLL_DOUBLE, "t", "is the time to expiration in years.")
		})
	// Function Wizard help
	.FunctionHelp("Returns the Black-Scholes-Merton put option value.")
	// Function Wizard category
	.Category("Financial")
	// URL linked to `Help on this function`
	.HelpTopic("https://en.wikipedia.org/wiki/Black%E2%80%93Scholes_model")
	.Documentation(R"xyzyx(
The Black-Scholes-Merton put option pricing formula calculates the theoretical price of a European put option.
)xyzyx")
);

// WINAPI calling convention must be specified
double WINAPI xll_bsm_put(double r, double S, double sigma, double K, double t)
{
#pragma XLLEXPORT // must be specified to export function
	return bsm_put(r, S, sigma, K, t);
}