package kz.dossier.extractor;

import kz.dossier.modelsDossier.MvFl;
import org.springframework.dao.DataAccessException;
import org.springframework.jdbc.core.ResultSetExtractor;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public class Mv_fl_extractor implements ResultSetExtractor<List<MvFl>> {
    @Override
    public List<MvFl> extractData(ResultSet rs) throws SQLException, DataAccessException {
        List<MvFl> data = new ArrayList<>();
        while (rs.next()) {
            MvFl res = new MvFl();
            res.setIin(rs.getString("iin"));
            res.setFirst_name(rs.getString("first_name"));
            res.setLast_name(rs.getString("last_name"));
            res.setPatronymic(rs.getString("patronymic"));
            data.add(res);
        }
        return data;
    }
}
