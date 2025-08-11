use xladd_derive::xl_func;
use thiserror::Error;

#[derive(Debug, Clone)]
pub struct OptionParameters {
    pub share_price: PositiveFloat,
    pub strike_price: PositiveFloat,
    pub time_to_maturity: PositiveFloat,
    pub vesting_period: PositiveFloat,
    pub risk_free: Rate,
    pub sigma: Volatility,
    pub div_rate: Rate,
    pub exit_pre_vesting: Rate,
    pub exit_post_vesting: Rate,
    pub multiple: PositiveFloat,
    pub steps: PositiveInt,
}

#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct PositiveFloat(pub f64);

#[derive(Debug, Clone)]
pub struct PositiveInt(pub usize);

#[derive(Debug, Clone)]
pub struct Rate(pub f64);

#[derive(Debug, Clone)]
pub struct Volatility(pub f64);

impl OptionParameters {
    pub fn new(
        share_price: f64,
        strike_price: f64,
        time_to_maturity: f64,
        vesting_period: f64,
        risk_free: f64,
        sigma: f64,
        div_rate: f64,
        exit_pre_vesting: f64,
        exit_post_vesting: f64,
        multiple: f64,
        steps: usize,
    ) -> Result<Self, ParameterError> {
        Ok(Self {
            share_price: PositiveFloat::new(share_price, "share_price")?,
            strike_price: PositiveFloat::new(strike_price, "strike_price")?,
            time_to_maturity: PositiveFloat::new(time_to_maturity, "time_to_maturity")?,
            vesting_period: PositiveFloat::new(vesting_period, "vesting_period")?.min_f64(time_to_maturity)?,
            risk_free: Rate::new(risk_free, "risk_free")?,
            sigma: Volatility::new(sigma)?,
            div_rate: Rate::new(div_rate, "div_rate")?,
            exit_pre_vesting: Rate::new(exit_pre_vesting, "exit_pre_vesting")?,
            exit_post_vesting: Rate::new(exit_post_vesting, "exit_post_vesting")?,
            multiple: PositiveFloat::new(multiple, "multiple")?,
            steps: PositiveInt::new(steps, "steps")?,
        })
    }
}

#[derive(Error, Debug)]
pub enum ParameterError {
    #[error("{parameter} must be positive, got {value}")]
    InvalidPositiveValue { parameter: &'static str, value: f64 },
    
    #[error("{parameter} must be positive, got {value}")]
    InvalidPositiveInt { parameter: &'static str, value: usize },
    
    #[error("Volatility must be non-negative, got {value}")]
    InvalidVolatility {value: f64 },
    
    #[error("{parameter} must be between 0 and 1, got {value}")]
    InvalidRate { parameter: &'static str, value: f64 },
}

impl PositiveFloat {
    /// Creates a new PositiveFloat if the value is positive and finite
    pub fn new(value: f64, parameter_name: &'static str) -> Result<Self, ParameterError> {
        if value >= 0.0 && value.is_finite() {
            Ok(PositiveFloat(value))
        } else {
            Err(ParameterError::InvalidPositiveValue { 
                parameter: parameter_name, 
                value 
            })
        }
    }

    pub fn min(self, other: PositiveFloat) -> PositiveFloat {
        if self.0 <= other.0 {
            self
        } else {
            other
        }
    }
    pub fn min_f64(self, other: f64) -> Result<PositiveFloat, ParameterError> {
        let other_positive = PositiveFloat::new(other, "comparison_value")?;
        Ok(self.min(other_positive))
    }
}

impl PositiveInt {
    /// Creates a new PositiveFloat if the value is positive and finite
    pub fn new(value: usize, parameter_name: &'static str) -> Result<Self, ParameterError> {
        if value > 0 {
            Ok(PositiveInt(value))
        } else {
            Err(ParameterError::InvalidPositiveInt { 
                parameter: parameter_name, 
                value 
            })
        }
    }
}

impl Rate {
    /// Creates a new Rate if the value is positive and finite
    pub fn new(value: f64, parameter_name: &'static str) -> Result<Self, ParameterError> {
        if value > 0.0 && value.is_finite() && value < 1.0{
            Ok(Rate(value))
        } else {
            Err(ParameterError::InvalidRate { 
                parameter: parameter_name, 
                value 
            })
        }
    }
}

impl Volatility {
    /// Creates a new Rate if the value is positive and finite
    pub fn new(value: f64) -> Result<Self, ParameterError> {
        if value > 0.0 && value.is_finite() {
            Ok(Volatility(value))
        } else {
            Err(ParameterError::InvalidVolatility { 
                value,
            })
        }
    }
}

/// # Description
/// Black-Scholes call option value for European options
/// # Arguments
/// * `share_price`: Current share price
/// * `strike_price`: Strike price of the option
/// * `time_to_maturity` - Time to maturity in years
/// * `risk_free` - Risk-free interest rate
/// * `div_rate` - Dividend yield
/// * `sigma` - Volatility
/// 
/// # Returns
/// Call option value using Black-Scholes formula
pub fn black_scholes_call_option_value(
    share_price: f64,
    strike_price: f64,
    time_to_maturity: f64,
    risk_free: f64,
    div_rate: f64,
    sigma: f64,
) -> f64 {
    // Handle zero strike price case
    let strike_price = if strike_price == 0.0 { 0.001 } else { strike_price };
    
    if sigma != 0.0 {
        // Standard Black-Scholes formula
        let d1 = (share_price / strike_price).ln() 
            + time_to_maturity * (risk_free - div_rate + 0.5 * sigma * sigma);
        let d1 = d1 / (sigma * time_to_maturity.sqrt());
        
        let d2 = d1 - sigma * time_to_maturity.sqrt();
        
        let call_option = share_price * (-div_rate * time_to_maturity).exp() * normal_cdf(d1)
            - (-risk_free * time_to_maturity).exp() * strike_price * normal_cdf(d2);
        
        call_option
    } else {
        // Zero volatility case - deterministic payoff
        let discounted_share_price = share_price * (-div_rate * time_to_maturity).exp();
        let discounted_strike_price = strike_price * (-risk_free * time_to_maturity).exp();
        
        (discounted_share_price - discounted_strike_price).max(0.0)
    }
}

/// Computes the cumulative distribution function (CDF) of the standard normal distribution.
///
/// Uses the Abramowitz and Stegun approximation (formula 7.1.26) for numerical accuracy.
///
/// # Arguments
///
/// * `x` - The input value for which to compute the CDF.
///
/// # Returns
///
/// The cumulative probability that a standard normal variable is less than or equal to `x`.
///
/// # Examples
///
/// ```
/// use my_crate::normal_cdf;
/// let p = normal_cdf(0.0);
/// assert!((p - 0.5).abs() < 1e-6);
/// ```
pub fn normal_cdf(x: f64) -> f64 {
    const A1: f64 = 0.254829592;
    const A2: f64 = -0.284496736;
    const A3: f64 = 1.421413741;
    const A4: f64 = -1.453152027;
    const A5: f64 = 1.061405429;
    const P: f64 = 0.3275911;
    
    let sign = if x < 0.0 { -1.0 } else { 1.0 };
    let x = x.abs() / std::f64::consts::SQRT_2;
    
    // A&S formula 7.1.26
    let t = 1.0 / (1.0 + P * x);
    let y = 1.0 - (((((A5 * t + A4) * t) + A3) * t + A2) * t + A1) * t * (-x * x).exp();
    
    0.5 * (1.0 + sign * y)
}

/// Computes the value of an employee stock option using a binomial tree model.
///
/// The function accounts for early exercise behavior, vesting periods, and 
/// employee exit probabilities before and after vesting. It uses a 
/// discrete-time binomial tree to model the option value over a given number 
/// of time steps, incorporating dividends and risk-neutral valuation.
///
/// # Parameters
///
/// * `share_price`-Current price of the underlying share.
/// * `strike_price` -Strike (exercise) price of the option.
/// * `maturity`- Total time to maturity of the option, in years.
/// * `vesting_period`:Vesting period during which the option cannot be exercised, in years.
/// * `risk_free` :Annualized risk-free interest rate (as a decimal, e.g., 0.07 for 7%).
/// * `sigma`: Volatility of the underlying share (as a decimal).
/// * `div_rate` - Dividend yield of the share (as a decimal).
/// * `exit_pre_vesting`:Annualized probability of employee exit before vesting (as a decimal).
/// * `exit_post_vesting` - Annualized probability of employee exit after vesting (as a decimal).
/// * `multiple` - Payoff multiplier applied to the intrinsic value (e.g. 1.0 for standard options).
/// * `steps` - Number of time steps in the binomial tree.
///
/// # Returns
///
/// A f64 vector of the estimated present value and Macaulay duration of an employee stock option
///
/// # Examples
///
/// ```
/// let value = binomial_option_value(
///     100.0,    // share_price
///     100.0,    // strike_price
///     5.0,      // maturity
///     2.0,      // vesting_period
///     0.05,     // risk_free
///     0.3,      // sigma
///     0.02,     // div_rate
///     0.1,      // exit_pre_vesting
///     0.05,     // exit_post_vesting
///     1.0,      // multiple
///     100       // steps
/// );
/// assert!(value > 0.0);
/// ```
#[xl_func()]
pub fn binomial_option_value(
    share_price: f64,
    strike_price: f64,
    time_to_maturity: f64,
    vesting_period: f64,
    risk_free: f64,
    sigma: f64,
    div_rate: f64,
    exit_pre_vesting: f64,
    exit_post_vesting: f64,
    multiple: f64,
    steps: i32,
) -> Result<Vec<f64>, Box<dyn std::error::Error>> {

    // Input validation and adjustments
    let steps = steps as usize;
    let strike_price = if strike_price == 0.0 { 0.001 } else { strike_price };
    let vesting_period = vesting_period.min(time_to_maturity);
    
    // Early exit for zero maturity
    if time_to_maturity == 0.0 {
        return Ok(vec![(share_price - strike_price).max(0.0), 0.0]);
    };
    
    // European option shortcut if vesting equals maturity
    if (vesting_period - time_to_maturity).abs() < f64::EPSILON {
        return Ok(vec![black_scholes_call_option_value(
                        share_price, strike_price, time_to_maturity, risk_free,
                        div_rate, sigma,),
                    time_to_maturity,]);
    }
    
    // Binomial tree parameters
    let dt = time_to_maturity / steps as f64;
    let u = (sigma * dt.sqrt()).exp();
    let d = 1.0 / u;
    let r = (risk_free * dt).exp();
    
    // Risk-neutral probability (handle zero sigma case)
    let p = if (u - d).abs() < f64::EPSILON {
        1.0
    } else {
        (((risk_free - div_rate) * dt).exp() - d) / (u - d)
    };
    
    // Vesting period in discrete time steps
    let vest_step = ((vesting_period / dt) + 0.001) as usize;
    
    // Exit probabilities per time step
    let px = (1.0 - exit_post_vesting).powf(dt);  // Prob of not exiting post-vesting
    let qx = 1.0 - px;                            // Prob of exiting post-vesting
    let px_pre = (1.0 - exit_pre_vesting).powf(dt); // Prob of not exiting pre-vesting
    
    // Pre-compute u and d powers for efficiency
    let u_powers: Vec<f64> = (0..=steps).map(|i| u.powi(i as i32)).collect();
    let d_powers: Vec<f64> = (0..=steps).map(|i| d.powi(i as i32)).collect();
    
    // Initialize matrices using flat arrays for better cache locality
    let matrix_size = (steps + 1) * (steps + 1);
    let mut share_price_matrix = vec![0.0; matrix_size];
    let mut intrinsic_value = vec![0.0; matrix_size];
    let mut option_value = vec![0.0; matrix_size];
    let mut macaulay_denominator = vec![0.0; matrix_size];
    let mut macaulay_numerator = vec![0.0; matrix_size];
    
    // Helper closure for 2D indexing into flat arrays
    let idx = |i: usize, j: usize| i * (steps + 1) + j;
    
    // Calculate share prices and intrinsic values at each node
    for i in (0..=steps).rev() {
        for j in 0..=i {
            share_price_matrix[idx(i, j)] = share_price * u_powers[j] * d_powers[i - j];
            intrinsic_value[idx(i, j)] = (share_price_matrix[idx(i, j)] - strike_price).max(0.0);
        }
    }
    
    // Initialize terminal conditions at maturity
    for i in 0..=steps {
        option_value[idx(steps, i)] = intrinsic_value[idx(steps, i)];
        macaulay_denominator[idx(steps, i)] = intrinsic_value[idx(steps, i)];
        macaulay_numerator[idx(steps, i)] = intrinsic_value[idx(steps, i)] * time_to_maturity;
    }
    
    // Backward induction through the binomial tree
    for i in (0..steps).rev() {
        for j in 0..=i {
            let pv_option_one_period = 
                (p * option_value[idx(i + 1, j + 1)] + (1.0 - p) * option_value[idx(i + 1, j)]) / r;
            
            if i >= vest_step {
                // Post-vesting period: optimal exercise or multiple trigger
                let should_exercise = intrinsic_value[idx(i, j)] > pv_option_one_period
                    || share_price_matrix[idx(i, j)] >= strike_price * multiple;
                
                if should_exercise {
                    option_value[idx(i, j)] = intrinsic_value[idx(i, j)];
                    macaulay_denominator[idx(i, j)] = intrinsic_value[idx(i, j)];
                    macaulay_numerator[idx(i, j)] = intrinsic_value[idx(i, j)] * (i as f64) * dt;
                } else {
                    option_value[idx(i, j)] = px * pv_option_one_period + qx * intrinsic_value[idx(i, j)];
                    
                    macaulay_denominator[idx(i, j)] = px * (
                        p * macaulay_denominator[idx(i + 1, j + 1)]
                        + (1.0 - p) * macaulay_denominator[idx(i + 1, j)]
                    ) + (1.0 - px) * intrinsic_value[idx(i, j)];
                    
                    macaulay_numerator[idx(i, j)] = px * (
                        p * macaulay_numerator[idx(i + 1, j + 1)]
                        + (1.0 - p) * macaulay_numerator[idx(i + 1, j)]
                    ) + (1.0 - px) * intrinsic_value[idx(i, j)] * (i as f64) * dt;
                }
            } else {
                // Pre-vesting period: cannot exercise, only exit rate applies
                option_value[idx(i, j)] = px_pre * pv_option_one_period;
                
                macaulay_denominator[idx(i, j)] = 
                    p * macaulay_denominator[idx(i + 1, j + 1)]
                    + (1.0 - p) * macaulay_denominator[idx(i + 1, j)];
                    
                macaulay_numerator[idx(i, j)] = 
                    p * macaulay_numerator[idx(i + 1, j + 1)]
                    + (1.0 - p) * macaulay_numerator[idx(i + 1, j)];
            }
        }
    }
    
    // Calculate expected life using Macaulay duration approach
    let expected_life = if macaulay_denominator[idx(0, 0)] != 0.0 {
        macaulay_numerator[idx(0, 0)] / macaulay_denominator[idx(0, 0)]
    } else {
        0.0
    };
    
    Ok(vec![option_value[idx(0, 0)], expected_life])
}


#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_zero_maturity() {
        let result = binomial_option_value(
            100.0, // share_price
            90.0,  // strike_price
            0.0,   // maturity
            0.0,   // vesting_period
            0.05,  // risk_free
            0.3,   // sigma
            0.0,   // div_rate
            0.1,   // exit_pre_vesting
            0.1,   // exit_post_vesting
            2.0,   // multiple
            100,   // steps
        ).unwrap();
        
        assert_eq!(result[0], 10.0); // 100 - 90
        assert_eq!(result[1], 0.0);
    }

    #[test]
    fn test_basic_option_value() {
        let result = binomial_option_value(
            100.0, // share_price
            90.0,  // strike_price
            1.0,   // maturity
            0.25,  // vesting_period (3 months)
            0.05,  // risk_free
            0.3,   // sigma
            0.0,   // div_rate
            0.1,   // exit_pre_vesting
            0.1,   // exit_post_vesting
            2.0,   // multiple
            100,   // steps
        ).unwrap();
        
        assert!(result[0] > 0.0);
        assert!(result[1] > 0.0);
        assert!(result[1] <= 1.0);
    }
}
