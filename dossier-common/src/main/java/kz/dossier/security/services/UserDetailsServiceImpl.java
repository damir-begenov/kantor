package kz.dossier.security.services;

import kz.dossier.security.models.User;
import kz.dossier.security.repository.UserRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.core.userdetails.UserDetailsService;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.security.Principal;

@Service
public class UserDetailsServiceImpl implements UserDetailsService {
  @Autowired
  UserRepository userRepository;

  @Override
  @Transactional
  public UserDetails loadUserByUsername(String username) throws UsernameNotFoundException {
    User user = userRepository.findByUsername(username)
        .orElseThrow(() -> new UsernameNotFoundException("User Not Found with username: " + username));

    return UserDetailsImpl.build(user);
  }
  @Transactional
  public User loadUserByUsernamek(Principal principal) throws UsernameNotFoundException {
    User user = userRepository.findByUsername(principal.getName())
            .orElseThrow(() -> new UsernameNotFoundException("User Not Found with username: " + principal.getName()));
    return user;
  }


}
