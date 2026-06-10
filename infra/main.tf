# =============================================================
# infra/main.tf
# Provisions a local kind cluster.
# State is stored in HCP Terraform (free tier) — see README
# for one-time setup instructions.
# =============================================================

terraform {
  required_version = ">= 1.0"

  required_providers {
    kind = {
      source  = "tehcyx/kind"
      version = "~> 0.4"
    }
  }

  # HCP Terraform free-tier backend
  # Replace <YOUR_ORG> with your HCP Terraform organisation name
  # after completing the one-time setup at: https://app.terraform.io
  cloud {
    organization = "<YOUR_ORG>"
    workspaces {
      name = "gitops-platform-demo"
    }
  }
}

provider "kind" {}

# ── kind cluster ───────────────────────────────────
resource "kind_cluster" "gitops" {
  name            = "gitops-demo"
  wait_for_ready  = true

  kind_config {
    kind        = "Cluster"
    api_version = "kind.x-k8s.io/v1alpha4"

    node {
      role = "control-plane"
    }

    node {
      role = "worker"
    }

    node {
      role = "worker"
    }
  }
}

# ── outputs ────────────────────────────────────────
output "cluster_name" {
  description = "kind cluster name"
  value       = kind_cluster.gitops.name
}

output "kubeconfig" {
  description = "kubeconfig for the cluster"
  value       = kind_cluster.gitops.kubeconfig
  sensitive   = true
}
