package kz.dossier.security.models;

import jakarta.annotation.Nullable;
import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Entity
@Table(name = "roles")
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class Role {
  @Id
  @GeneratedValue(strategy = GenerationType.IDENTITY)
  private Integer id;

  @Column(length = 20)
  @Enumerated(EnumType.STRING)
  private ERole name;
  @ElementCollection
  private List<String> relations;

  @ElementCollection
  private List<String> person_properties;
  @ElementCollection
  private List<String> company_properties;
}