# Repository Governance

Pi-Office is a public, source-available repository under the
[`Pi-Office Personal Non-Commercial License v1.0`](../LICENSE).

## License Posture

- Personal, educational, research, and evaluation forks are allowed.
- Commercial use, hosting, paid distribution, paid support, or commercial
  integration requires prior written permission from the copyright holder.
- Public forks or derivatives must credit Pi-Office and link to the original
  repository: <https://github.com/cryptekbits/Pi-Office>.
- This is a non-commercial source-available license, not an OSI-approved
  open-source license.

## Main Branch Protection

The `main` branch is protected on GitHub with a repository ruleset named
`Protect main`.

Rules currently enforced:

- Pull requests are required before merging into `main`.
- One approving review is required.
- Code-owner review is required for files matched by `.github/CODEOWNERS`.
- Stale approvals are dismissed when new commits are pushed.
- The most recent push must be approved by someone other than the pusher.
- Review conversations must be resolved before merge.
- Force pushes and branch deletion are blocked.

GitHub personal repositories do not support named-user push restrictions in
classic branch protection, and repository rulesets do not support named-user
bypass actors. Pi-Office therefore uses the repository-admin role as the
direct-push bypass actor. At the time this policy was applied, the only
visible collaborator/admin was `cryptekbits`.

To allow additional people to approve protected PRs, add them to
`.github/CODEOWNERS` and give them appropriate repository access. To allow
additional people to bypass PR-only updates and push directly to `main`,
they must be granted repository admin access or the repository must move to
an organization where team/user restrictions can be modeled more precisely.
