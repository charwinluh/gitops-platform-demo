# gitops-platform-demo

Demo of a GitOps pipeline: Terraform provisions a local kind cluster; the intended stack on top is ArgoCD, Prometheus and Grafana with SLOs (per the README). Early stage: only the Terraform cluster definition exists.

## Layout
- `infra/main.tf`: Terraform config. Uses the `tehcyx/kind` provider (`~> 0.4`), Terraform `>= 1.0`, and an HCP Terraform `cloud` backend (workspace `gitops-platform-demo`). Creates a kind cluster `gitops-demo` with one control-plane and two worker nodes. Outputs `cluster_name` and a sensitive `kubeconfig`.
- `app/`, `k8s/`, `monitoring/`: empty placeholders (`.gitkeep`) for the application, Kubernetes manifests and monitoring config.
- `README.md`: one-line description.

## Working on it
- The `organization` value in the `cloud` block is the placeholder `<YOUR_ORG>`; set it to your own HCP Terraform organization before `terraform init`.
- Needs Docker and the kind tooling to create the cluster locally.
- Typical flow: `terraform -chdir=infra init`, `plan`, `apply`. These were not run when writing this file.
- Keep new manifests under `k8s/`, application code under `app/`, dashboards and alert rules under `monitoring/`.

## Rules
- Never commit Terraform state, `*.tfvars` (except `*.tfvars.example`), `.terraform/`, lock overrides, kubeconfigs or credentials; `.gitignore` already excludes state and tfvars.
- Mark secrets as `sensitive` in Terraform outputs, as the existing kubeconfig output does.
- Run `terraform fmt` and `terraform validate` before committing infra changes.
