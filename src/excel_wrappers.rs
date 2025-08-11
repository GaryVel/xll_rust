use xladd_derive::xl_func;
use crate::actuarial::option_pricing;
use crate::actuarial::option_pricing::{OptionParameters, PositiveFloat, PositiveInt, Rate, Volatility};

// use ndarray::Array2;

/// Amazing function to calculate the sum of three numbers
#[xl_func(
    category = "Math",
    params(
        num1 = "First value to addGV",
        different_number = "Second value to addGV",
        a_new_number = "Third value to addGV"
    )
)]
fn add_xx2(
    num1: f64,
    different_number: f64,
    a_new_number: f64
) -> Result<f64, Box<dyn std::error::Error>> {
    Ok(num1 + different_number + a_new_number)
}

// Custom category
#[xl_func(category="Math & Trig")]
fn my_math_func(x: f64) -> f64 { x }   // note that this returns just f64 and not a Result. Not recommended in majority of cases

// Custom prefix (default: "xl")
#[xl_func(prefix="my")]
fn calc_value(x: f64) -> Result<f64, Box<dyn std::error::Error>> { Ok(x) }

// Custom Excel function name
#[xl_func(rename="CustomName")]
fn some_function(x: f64) -> Result<f64, Box<dyn std::error::Error>> { Ok(x) }

// Thread-safe function (adds $ to registration string)
#[xl_func(threadsafe)]
fn thread_safe_func(x: f64) -> Result<f64, Box<dyn std::error::Error>> { Ok(x) }

// Single-threaded (default behavior)
#[xl_func(single_threaded)]
fn single_thread_func(x: f64) -> Result<f64, Box<dyn std::error::Error>> { Ok(x) }

// Combine multiple options
#[xl_func(category="Financial", prefix="fin", threadsafe)]
fn advanced_calc(rate: f64, years: f64) -> Result<f64, Box<dyn std::error::Error>> { 
    Ok(rate * years) 
}

/// AF function to calculate value of optimal option (participant exercises at most optimal time)
/// #Parameters
/// * share_price: share price at grant date
/// * strike_price: price at which option is to be exercised
/// * time_to_maturity: term to maturity in years
/// * vesting_period: term until end of vesting period in years (vesting period <= term to maturity)
/// * risk_free: risk free rate at the appropriate duration
/// * sigma: share volatility at the appropriate duration
/// * divrate: dividend rate
/// * exit_pre_vesting: exit rate before vesting date
/// * exit_post_vesting: exit rate after vesting date
/// * n: number of iterations to estimate value (1000 is plenty; 100 can work too)
#[xl_func()]
fn option_value_optimal(
    share_price: f64,
    strike_price: f64,
    time_to_maturity: f64,
    vesting_period: f64,
    risk_free: f64,
    sigma: f64,
    divrate: f64,
    exit_pre_vesting: f64,
    exit_post_vesting: f64,
    n: f64,
) -> Result<Vec<f64>, Box<dyn std::error::Error>> {
    let multiple: f64 = 1e7;
    let params = OptionParameters {
        share_price: PositiveFloat(share_price),
        strike_price: PositiveFloat(strike_price),
        time_to_maturity: PositiveFloat(time_to_maturity),
        vesting_period: PositiveFloat(vesting_period),
        risk_free: Rate(risk_free),
        sigma: Volatility(sigma),
        div_rate: Rate(divrate),
        exit_pre_vesting: Rate(exit_pre_vesting),
        exit_post_vesting: Rate(exit_post_vesting),
        multiple: PositiveFloat(multiple),
        steps: PositiveInt(n as usize), // Assuming n is a positive integer
    };
    
    let result = option_pricing::binomial_option_value(
        share_price, strike_price, time_to_maturity, vesting_period,
        risk_free, sigma, divrate,
        exit_pre_vesting, exit_post_vesting,
        multiple,
        n as i32)?;
    Ok(result)
}

/// AF function to calculate value of non optimal option (participant exercises at most optimal time)
/// #Parameters
/// * share_price: share price at grant date
/// * strike_price: price at which option is to be exercised
/// * t: term to maturity in years
/// * vesting_period: term until end of vesting period in years (vesting period <= term to maturity)
/// * risk_free: risk free rate at the appropriate duration
/// * sigma: share volatility at the appropriate duration
/// * divrate: dividend rate
/// * exit_pre_vesting: exit rate before vesting date
/// * exit_post_vesting: exit rate after vesting date
/// * multiple: multiple of the strike price at which option holder assumed to exercise
/// * n: number of iterations to estimate value (1000 is plenty; 100 can work too)
#[xl_func()]
fn option_value_non_optimal(
    share_price: f64,
    strike_price: f64,
    t: f64,
    vesting_period: f64,
    risk_free: f64,
    sigma: f64,
    divrate: f64,
    exit_pre_vesting: f64,
    exit_post_vesting: f64,
    multiple: f64,
    n: f64,
) -> Result<Vec<f64>, Box<dyn std::error::Error>> {
    let result = option_pricing::binomial_option_value(
        share_price, strike_price, t, vesting_period,
        risk_free, sigma, divrate,
        exit_pre_vesting, exit_post_vesting,
        multiple,
        n as i32)?;
    Ok(result)
}
