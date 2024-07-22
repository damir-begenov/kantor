package kz.dossier.controller;

import jakarta.validation.Valid;
import kz.dossier.payload.request.LoginRequest;
import kz.dossier.payload.request.SignupRequest;
import kz.dossier.payload.response.JwtResponse;
import kz.dossier.payload.response.MessageResponse;
import kz.dossier.security.jwt.JwtUtils;
import kz.dossier.security.models.ERole;
import kz.dossier.security.models.Role;
import kz.dossier.security.models.User;
import kz.dossier.security.repository.RoleRepository;
import kz.dossier.security.repository.UserRepository;
import kz.dossier.security.services.UserDetailsImpl;
import kz.dossier.security.services.UserDetailsServiceImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.security.authentication.AuthenticationManager;
import org.springframework.security.authentication.UsernamePasswordAuthenticationToken;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.web.bind.annotation.*;

import java.security.Principal;
import java.util.HashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;

@CrossOrigin(origins = "*", maxAge = 3000)
@RestController
@RequestMapping("/api/pandora/auth")
public class AuthController {
  @Autowired
  AuthenticationManager authenticationManager;

  @Autowired
  UserRepository userRepository;
  @Autowired

  UserDetailsServiceImpl userDetailsService;
  @Autowired
  RoleRepository roleRepository;

  @Autowired
  PasswordEncoder encoder;

  @Autowired
  JwtUtils jwtUtils;

  @PostMapping("/signin")
  public ResponseEntity<?> authenticateUser(@Valid @RequestBody LoginRequest loginRequest) {
    Authentication authentication = authenticationManager.authenticate(
            new UsernamePasswordAuthenticationToken(loginRequest.getUsername(), loginRequest.getPassword()));
    SecurityContextHolder.getContext().setAuthentication(authentication);
    String jwt = jwtUtils.generateJwtToken(authentication);

    UserDetailsImpl userDetails = (UserDetailsImpl) authentication.getPrincipal();
    List<String> roles = userDetails.getAuthorities().stream()
            .map(item -> item.getAuthority())
            .collect(Collectors.toList());
    return ResponseEntity.ok(new JwtResponse(jwt,
            userDetails.getId(),
            userDetails.getUsername(),
            userDetails.getEmail(),
            roles));
  }
  @PostMapping("/changePassword")
  public void changePassword( @RequestParam String password, Principal principal, @RequestParam String username){
    Optional<User> user = userRepository.findByUsername(username);
    if (user.isPresent()) {
      user.get().setPassword(encoder.encode(password));
      System.out.println(user.get().getUsername());
      System.out.println(user.get().getEmail());
      System.out.println(password);
      userRepository.save(user.get());
    }
  }
  @PostMapping("/signup")
  public ResponseEntity<?> registerUser(@Valid @RequestBody SignupRequest signUpRequest) {

    if (userRepository.existsByUsername(signUpRequest.getUsername())) {
      return ResponseEntity
              .badRequest()
              .body(new MessageResponse("Error: Username is already taken!"));
    }
    if (userRepository.existsByEmail(signUpRequest.getEmail())) {
      return ResponseEntity
              .badRequest()
              .body(new MessageResponse("Error: Email is already in use!"));
    }

    User user = new User(signUpRequest.getEmail(),
            signUpRequest.getUsername(),
            encoder.encode(signUpRequest.getPassword()));

    Set<String> strRoles = signUpRequest.getRole();
    Set<Role> roles = new HashSet<>();

    if (signUpRequest.getLevel().equals("2")) {
      Role userRole = roleRepository.findByName(ERole.LEVEL_2_USER)
              .orElseThrow(() -> new RuntimeException("Error: Role is not found."));
      roles.add(userRole);  }
    if (signUpRequest.getLevel().equals("1")) {
      Role userRole = roleRepository.findByName(ERole.LEVEL_1_USER)
              .orElseThrow(() -> new RuntimeException("Error: Role is not found."));
      roles.add(userRole);  }if (signUpRequest.getLevel().equals("3")) {
      Role userRole = roleRepository.findByName(ERole.LEVEL_3_USER)
              .orElseThrow(() -> new RuntimeException("Error: Role is not found."));
      roles.add(userRole);  }if (signUpRequest.getLevel().equals("vip")) {
      Role userRole = roleRepository.findByName(ERole.VIP)
              .orElseThrow(() -> new RuntimeException("Error: Role is not found."));
      roles.add(userRole);  }if (signUpRequest.getLevel().equals("admin")) {
      Role userRole = roleRepository.findByName(ERole.ADMIN)
              .orElseThrow(() -> new RuntimeException("Error: Role is not found."));
      roles.add(userRole);  }
    user.setActive(true);
    user.setRoles(roles);
    userRepository.save(user);
    return ResponseEntity.ok(new MessageResponse("User registered successfully!"));
  }
}
