package kz.dossier.payload.request;

import jakarta.persistence.Lob;
import jakarta.validation.constraints.Size;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.Set;
@AllArgsConstructor
@NoArgsConstructor
@Getter
@Setter
public class SignupRequest {
  @Size(min = 1, max = 50)
  private String username;

  @Size(max = 50)
  private String email;

  private Set<String> role;

  @Size(min = 1, max = 40)
  private String password;
  //  @NotBlank
  private String level;
  @Lob
  private String user_photo;


}
