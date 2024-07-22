package kz.dossier.security.jwt;

import io.jsonwebtoken.*;
import io.jsonwebtoken.security.Keys;
import kz.dossier.security.services.UserDetailsImpl;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.security.core.Authentication;
import org.springframework.stereotype.Component;

import javax.crypto.SecretKey;
import java.util.Date;

@Component
public class JwtUtils {
  private static final Logger logger = LoggerFactory.getLogger(JwtUtils.class);

  @Value("Y2Fubm90cm9hbF9leGFtcGxlX2tleV9zdHJpbmc6Y2Fubm90cm9hbF9leGFtcGxlX2tleV9zdHJpbmc6Y2Fubm90cm9hbF9leGFtcGxlX2tleV9zdHJpbmc6Y2Fubm90cm9hbF9leGFtcGxlX2tleV9zdHJpbmc=")
  private String jwtSecret;

  @Value("89000000")
  private int jwtExpirationMs;

  public String generateJwtToken(Authentication authentication) {
    UserDetailsImpl userPrincipal = (UserDetailsImpl) authentication.getPrincipal();

    return Jwts.builder()
            .setSubject(userPrincipal.getUsername())
            .setIssuedAt(new Date())
            .setExpiration(new Date((new Date()).getTime() + jwtExpirationMs))
            .signWith(getSigningKey(), SignatureAlgorithm.HS256)
            .compact();
  }

  private SecretKey getSigningKey() {
    // Ensure that jwtSecret is long enough for the HS512 algorithm.
    byte[] keyBytes = jwtSecret.getBytes();
    if (keyBytes.length < 64) { // HS512 requires at least 64 bytes (512 bits)
      throw new IllegalArgumentException("JWT secret key must be at least 64 bytes for HS512 algorithm");
    }
    return Keys.hmacShaKeyFor(keyBytes);
  }

  public String getUserNameFromJwtToken(String token) {
    Claims claims = Jwts.parser()
            .setSigningKey(getSigningKey())
            .build()
            .parseClaimsJws(token)
            .getBody();

    return claims.getSubject();
  }

  public boolean validateJwtToken(String authToken) {
    try {
      Jwts.parser().setSigningKey(getSigningKey())
              .build()
              .parseClaimsJws(authToken);
      return true;
    } catch (SignatureException e) {
      logger.error("Invalid JWT signature: {}", e.getMessage());
    } catch (MalformedJwtException e) {
      logger.error("Invalid JWT token: {}", e.getMessage());
    } catch (ExpiredJwtException e) {
      logger.error("JWT token is expired: {}", e.getMessage());
    } catch (UnsupportedJwtException e) {
      logger.error("JWT token is unsupported: {}", e.getMessage());
    } catch (IllegalArgumentException e) {
      logger.error("JWT claims string is empty: {}", e.getMessage());
    }

    return false;
  }
}
