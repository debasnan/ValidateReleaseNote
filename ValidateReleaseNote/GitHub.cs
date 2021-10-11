using Octokit;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ValidateReleaseNote
{
    public class GitHub
    {
        public GitHub()
        {
        }

        public GitHubClient GetGitHubClient(string userName, string password, string token, string enterpriseUrl)
        {
            Credentials credentials = (!string.IsNullOrEmpty(token)) ? new Credentials(token, AuthenticationType.Bearer) : new Credentials(userName, password, AuthenticationType.Basic);

            var github = GetEnterpriseClient(new ProductHeaderValue("TestApp"), credentials, enterpriseUrl);

            return github;
        }

        public async Task<List<GitHubCommitFile>> GetAllFiles(GitHubClient github, string repositoryOwner, string repository, string gitReference)
        {

            var commitDetails = await github.Repository.Commit.Get(repositoryOwner, repository, gitReference);


            List<GitHubCommitFile> fileList = new List<GitHubCommitFile>();

            foreach (GitHubCommitFile file in commitDetails.Files)
            {
                fileList.Add(file);

            }

            return fileList;
        }

        private static GitHubClient GetEnterpriseClient(ProductHeaderValue productInformation,
           Credentials credentials, string enterpriseUrl)
        {
            var client = new GitHubClient(productInformation, new Uri(enterpriseUrl))
            {
                Credentials = credentials
            };

            return client;
        }
    }
}
