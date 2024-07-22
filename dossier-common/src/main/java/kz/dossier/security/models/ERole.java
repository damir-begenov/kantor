package kz.dossier.security.models;

import jakarta.persistence.Enumerated;

public enum ERole {
  ADMIN,
  VIP,
  LEVEL_1_USER,
  LEVEL_2_USER,
  LEVEL_3_USER
}
