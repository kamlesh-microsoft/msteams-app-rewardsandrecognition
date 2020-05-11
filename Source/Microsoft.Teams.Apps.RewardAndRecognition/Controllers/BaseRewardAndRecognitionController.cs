// <copyright file="BaseRewardAndRecognitionController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System.Linq;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    /// Base controller to handle user and company response API operations.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class BaseRewardAndRecognitionController : ControllerBase
    {
        /// <summary>
        /// Get claims of user.
        /// </summary>
        /// <returns>User claims.</returns>
        protected JwtClaims GetUserClaims()
        {
            var claims = this.User.Claims;
            var jwtClaims = new JwtClaims
            {
                FromId = claims.Where(claim => claim.Type == "http://schemas.microsoft.com/identity/claims/objectidentifier").Select(claim => claim.Value).First(),
                Upn = claims.Where(claim => claim.Type == "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn").Select(claim => claim.Value).First(),
            };

            return jwtClaims;
        }

        /// <summary>
        /// Creates the error response as per the status codes in case of error.
        /// </summary>
        /// <param name="statusCode">Describes the type of error.</param>
        /// <param name="errorMessage">Describes the error message.</param>
        /// <returns>Returns error response with appropriate message and status code.</returns>
        protected IActionResult GetErrorResponse(int statusCode, string errorMessage)
        {
            switch (statusCode)
            {
                case StatusCodes.Status401Unauthorized:
                    return this.StatusCode(
                        StatusCodes.Status401Unauthorized,
                        new ErrorResponse
                        {
                            StatusCode = "signinRequired",
                            ErrorMessage = errorMessage,
                        });
                case StatusCodes.Status400BadRequest:
                    return this.StatusCode(
                        StatusCodes.Status400BadRequest,
                        new ErrorResponse
                        {
                            StatusCode = "badRequest",
                            ErrorMessage = errorMessage,
                        });
                default:
                    return this.StatusCode(
                        StatusCodes.Status500InternalServerError,
                        new ErrorResponse
                        {
                            StatusCode = "internalServerError",
                            ErrorMessage = errorMessage,
                        });
            }
        }
    }
}